/*!
 * azure-maps-openlayers-dist.js
 *
 * A “dist” version (no imports, no Node) that registers:
 *    ol.source.AzureMaps
 *
 * Dependencies:
 *   • OpenLayers (dist/ol.js must be loaded first)
 *   • (Optional) ADAL.js (for authType==='aad')
 *
 * Usage:
 *   <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/ol@7.4.0/ol.css" />
 *   <script src="https://cdn.jsdelivr.net/npm/ol@7.4.0/dist/ol.js"></script>
 *   <!-- If using AAD auth: -->
 *   <script src="https://secure.aadcdn.microsoftonline-p.com/lib/1.0.18/js/adal.min.js"></script>
 *   <script src="azure-maps-openlayers-dist.js"></script>
 *
 *   <script>
 *     const azureSource = new ol.source.AzureMaps({
 *       tilesetId: 'microsoft.base.road',
 *       language: 'en-US',
 *       view: 'Auto',
 *       authOptions: {
 *         authType: 'subscriptionKey',
 *         subscriptionKey: 'YOUR_AZURE_MAPS_KEY',
 *         azMapsDomain: 'atlas.microsoft.com'
 *       }
 *     });
 *     const azureLayer = new ol.layer.Tile({ source: azureSource });
 *     const map = new ol.Map({
 *       target: 'map',
 *       layers: [azureLayer],
 *       view: new ol.View({
 *         center: [0, 0],
 *         zoom: 2,
 *         projection: 'EPSG:3857'
 *       })
 *     });
 *   </script>
 */
(function (ol) {
    'use strict';

    if (typeof ol === 'undefined') {
        throw new Error('OpenLayers (ol) must be loaded before azure-maps-openlayers-dist.js');
    }

    // ------------------------------
    // CONSTANTS & HELPERS
    // ------------------------------
    const Constants = {
        preferredCacheLocation: 'localStorage',
        storage: {
            accessTokenKey: 'access.token.key',
            testStorageKey: 'testStorage'
        },
        tokenExpiresIn: 3599,
        tokenRefreshClockSkew: 300, // 5 minutes
        errors: {
            tokenExpired: 'Token Expired, Try again'
        },
        AUTHORIZATION: 'authorization',
        AUTHORIZATION_SCHEME: 'Bearer',
        MAP_AGENT: 'Map-Agent',
        MS_AM_REQUEST_ORIGIN: 'Ms-Am-Request-Origin',
        MS_AM_REQUEST_ORIGIN_VALUE: 'MapControl',
        X_MS_CLIENT_ID: 'x-ms-client-id',
        SESSION_ID: 'Session-Id',
        SHORT_DOMAIN: 'atlas.microsoft.com',
        DEFAULT_DOMAIN: 'https://atlas.microsoft.com/',
        SDK_VERSION: '0.0.1',
        TARGET_SDK: 'OpenLayers',
        RENDERV2_VERSION: '2024-04-01',
    };

    class Helpers {
        /** Generates an RFC4122 v4 UUID (hex). */
        static uuid() {
            // From https://stackoverflow.com/a/2117523
            return ([1e7] + -1e3 + -4e3 + -8e3 + -1e11).replace(/[018]/g, c =>
                (
                    c ^
                    (crypto.getRandomValues(new Uint8Array(1))[0] & (15 >> (c / 4)))
                ).toString(16)
            );
        }
    }

    // ------------------------------
    // TIMERS
    // ------------------------------
    // A worker-based setTimeout that fires even in inactive tabs.
    const SetTimeoutWorkerCode = `
    onmessage = function (event) {
      var delay = event.data.time;
      setTimeout(() => {
        postMessage({ id: event.data.id });
      }, delay);
    };
  `;

    class Timers {
        static _workerTable = {};

        /** Schedules `callback` after `timeout` ms, even if the tab is inactive. */
        static setTimeout(callback, timeout) {
            const id = Math.round(Math.random() * 1e9);
            const blob = new Blob([SetTimeoutWorkerCode], { type: 'application/javascript' });
            const blobURL = URL.createObjectURL(blob);
            const worker = new Worker(blobURL);

            worker.addEventListener('message', e => {
                const w = Timers._workerTable[e.data.id];
                if (w) {
                    w.callback();
                    w.worker.terminate();
                    delete Timers._workerTable[e.data.id];
                }
            });

            Timers._workerTable[id] = { callback, worker };
            worker.postMessage({ id, time: timeout });
            return id;
        }

        /** Cancels a pending Timers.setTimeout. */
        static clearTimeout(id) {
            const w = Timers._workerTable[id];
            if (w) {
                w.worker.terminate();
                delete Timers._workerTable[id];
            }
        }
    }

    // ------------------------------
    // AUTHENTICATION MANAGER
    // ------------------------------
    /**
     * Manages Azure Maps authentication for three modes:
     *  • subscriptionKey → simply append ?subscription-key=… to each tile URL
     *  • aad             → uses ADAL.js (window.AuthenticationContext) to fetch/refresh tokens
     *  • anonymous       → user provides getToken(success,fail) to fetch tokens from their own backend
     */
    class AuthenticationManager {
        static instance = null;
        static defaultAuthContext = null;
        static sessionId = Helpers.uuid();
        static fallbackStorage = {};

        constructor(authOptions) {
            this.options = authOptions;
            this._initialized = false;
        }

        /** Compare two authOptions for equality (to reuse a singleton manager). */
        _compareOptions(authOptions) {
            const o = this.options;
            return (
                authOptions.azMapsDomain === o.azMapsDomain &&
                authOptions.aadAppId === o.aadAppId &&
                authOptions.aadInstance === o.aadInstance &&
                authOptions.aadTenant === o.aadTenant &&
                authOptions.authType === o.authType &&
                authOptions.clientId === o.clientId &&
                authOptions.getToken === o.getToken &&
                authOptions.subscriptionKey === o.subscriptionKey
            );
        }

        /**
         * Returns a singleton AuthenticationManager for the given options.
         * authOptions = { authType:'subscriptionKey'|'aad'|'anonymous', subscriptionKey, clientId, aadTenant, aadInstance, getToken, azMapsDomain }
         */
        static getInstance(authOptions) {
            if (authOptions && authOptions.authType) {
                let domain = authOptions.azMapsDomain || '';
                if (/^\w+:\/\//.test(domain)) {
                    authOptions.azMapsDomain = domain.replace(/^\w+:\/\//, '');
                }
                if (
                    AuthenticationManager.instance &&
                    AuthenticationManager.instance._compareOptions(authOptions)
                ) {
                    return AuthenticationManager.instance;
                }
                const mgr = new AuthenticationManager(authOptions);
                if (!AuthenticationManager.instance) {
                    AuthenticationManager.instance = mgr;
                }
                return mgr;
            }
            if (AuthenticationManager.instance) {
                return AuthenticationManager.instance;
            }
            throw new Error('Azure Maps credentials not specified.');
        }

        /** True if initialize() has completed successfully. */
        isInitialized() {
            return this._initialized;
        }

        /**
         * Initialize authentication; returns a Promise that resolves once tokens (if any) are ready.
         */
        initialize() {
            if (!this.initPromise) {
                this.initPromise = new Promise((resolve, reject) => {
                    const opt = this.options;

                    if (opt.authType === 'subscriptionKey') {
                        // Nothing to do for subscriptionKey mode.
                        this._initialized = true;
                        return resolve();
                    }

                    if (opt.authType === 'aad') {
                        // Create or reuse a default ADAL context
                        opt.authContext =
                            opt.authContext || AuthenticationManager._getDefaultAuthContext(opt);

                        // If the current window is an ADAL callback (hidden iframe), let ADAL handle it
                        opt.authContext.handleWindowCallback();
                        if (opt.authContext.getLoginError()) {
                            return reject(
                                new Error('Error logging in AAD user: ' + opt.authContext.getLoginError())
                            );
                        }
                        if (opt.authContext.isCallback(window.location.hash)) {
                            // ADAL is doing redirect/renew in this iframe—stop here.
                            return;
                        }

                        // Kick off login + token acquire on next tick (so user code can attach listeners first)
                        Timers.setTimeout(() => this._loginAndAcquire(resolve, reject), 0);
                        return;
                    }

                    if (opt.authType === 'anonymous') {
                        // Let user’s getToken fetch a token now
                        this._initialized = true;
                        return resolve(this._triggerTokenFetch());
                    }

                    reject(new Error('Invalid authentication type.'));
                });
            }
            return this.initPromise;
        }

        /**
         * If authType==='aad', create (or reuse) a shared ADAL AuthenticationContext.
         * Requires that ADAL.js has been loaded (global AuthenticationContext).
         */
        static _getDefaultAuthContext(options) {
            if (!options.aadAppId) {
                throw new Error('No AAD app ID was specified.');
            }
            if (!options.aadTenant) {
                throw new Error('No AAD tenant was specified.');
            }
            if (!AuthenticationManager.defaultAuthContext) {
                if (typeof AuthenticationContext === 'undefined') {
                    throw new Error(
                        'ADAL.js must be loaded before using authType==="aad" (no AuthenticationContext found).'
                    );
                }
                AuthenticationManager.defaultAuthContext = new AuthenticationContext({
                    instance: options.aadInstance || 'https://login.microsoftonline.com/',
                    tenant: options.aadTenant,
                    clientId: options.aadAppId,
                    cacheLocation: Constants.preferredCacheLocation
                });
            }
            return AuthenticationManager.defaultAuthContext;
        }

        /**
         * After ADAL login, acquire an Azure Maps token and resolve/reject accordingly.
         */
        _loginAndAcquire(resolve, reject) {
            const opt = this.options;
            const acquireAndResolve = () => {
                opt.authContext.acquireToken(Constants.DEFAULT_DOMAIN, error => {
                    if (error) {
                        reject(new Error(error));
                    } else {
                        this._initialized = true;
                        resolve();
                    }
                });
            };

            const cachedToken = opt.authContext.getCachedToken(opt.aadAppId);
            const cachedUser = opt.authContext.getCachedUser();
            if (cachedToken && cachedUser) {
                return acquireAndResolve();
            }

            if (!opt.authContext.loginInProgress()) {
                opt.authContext.login();
            }

            const poll = setInterval(() => {
                if (!opt.authContext.loginInProgress()) {
                    clearInterval(poll);
                    if (opt.authContext.getCachedToken(opt.aadAppId)) {
                        return acquireAndResolve();
                    } else {
                        reject(
                            new Error(
                                opt.authContext.getLoginError() ||
                                'AAD auth context is not logged in for AAD App ID: ' + opt.aadAppId
                            )
                        );
                    }
                }
            }, 25);
        }

        /**
         * Returns a valid token string or throws if something went wrong.
         *   • authType==='aad'       → fetch from ADAL cache or renew if expired
         *   • authType==='anonymous' → return locally cached token, or trigger getToken()
         *   • authType==='subscriptionKey' → return subscriptionKey
         */
        getToken() {
            const opt = this.options;
            if (opt.authType === 'aad') {
                let t = opt.authContext.getCachedToken(Constants.DEFAULT_DOMAIN);
                if (!t) {
                    if (!opt.authContext.getCachedUser()) {
                        opt.authContext.login();
                    }
                    opt.authContext.acquireToken(Constants.DEFAULT_DOMAIN, (error, renewed) => {
                        if (!error) {
                            t = renewed;
                        }
                    });
                }
                return t;
            }

            if (opt.authType === 'anonymous') {
                let token = this._getItem(Constants.storage.accessTokenKey);
                if (!token) {
                    this._triggerTokenFetch();
                } else {
                    const expiresIn = this._getTokenExpiry(token);
                    if (expiresIn < 300000 && expiresIn > 0) {
                        this._triggerTokenFetch();
                    } else if (expiresIn <= 0) {
                        this._saveItem(Constants.storage.accessTokenKey, '');
                        throw new Error(Constants.errors.tokenExpired);
                    }
                }
                return token;
            }

            if (opt.authType === 'subscriptionKey') {
                return opt.subscriptionKey;
            }

            throw new Error('Invalid authentication type.');
        }

        /**
         * Calls the user‐supplied getToken() (for authType==='anonymous').
         * Expects getToken(successFn(token), errorFn(err)). Caches token + schedules renewal.
         */
        _triggerTokenFetch() {
            return new Promise((resolve, reject) => {
                this.options.getToken(
                    token => {
                        try {
                            const timeout = this._getTokenExpiry(token) - Constants.tokenRefreshClockSkew;
                            this._storeAccessToken(token);
                            Timers.clearTimeout(this.tokenTimeOutHandle);
                            this.tokenTimeOutHandle = Timers.setTimeout(
                                () => this._triggerTokenFetch(),
                                timeout
                            );
                            resolve();
                        } catch {
                            reject(new Error('Invalid token returned by getToken()'));
                        }
                    },
                    error => {
                        reject(error);
                    }
                );
            });
        }

        /** Decode a JWT to compute ms until expiration (minus 5m skew). */
        _getTokenExpiry(token) {
            const payload = JSON.parse(atob(token.split('.')[1]));
            const expires = payload.exp * 1000; // exp is in seconds
            return expires - Date.now() - 300000;
        }

        /** Store token in local/session storage (or fallback object). */
        _storeAccessToken(token) {
            this._saveItem(Constants.storage.accessTokenKey, token);
        }

        /** Save a key/value to storage. */
        _saveItem(key, value) {
            if (this._supportsLocalStorage()) {
                localStorage.setItem(key, value);
                return true;
            } else if (this._supportsSessionStorage()) {
                sessionStorage.setItem(key, value);
                return true;
            } else {
                AuthenticationManager.fallbackStorage[key] = value;
                return true;
            }
        }

        /** Retrieve a key from storage. */
        _getItem(key) {
            if (this._supportsLocalStorage()) {
                return localStorage.getItem(key);
            } else if (this._supportsSessionStorage()) {
                return sessionStorage.getItem(key);
            } else {
                return AuthenticationManager.fallbackStorage[key];
            }
        }

        /** Test localStorage support. */
        _supportsLocalStorage() {
            try {
                const wls = window.localStorage;
                const testKey = Constants.storage.testStorageKey;
                if (!wls) return false;
                wls.setItem(testKey, 'A');
                if (wls.getItem(testKey) !== 'A') return false;
                wls.removeItem(testKey);
                return wls.getItem(testKey) === null;
            } catch {
                return false;
            }
        }

        /** Test sessionStorage support. */
        _supportsSessionStorage() {
            try {
                const wss = window.sessionStorage;
                const testKey = Constants.storage.testStorageKey;
                if (!wss) return false;
                wss.setItem(testKey, 'A');
                if (wss.getItem(testKey) !== 'A') return false;
                wss.removeItem(testKey);
                return wss.getItem(testKey) === null;
            } catch {
                return false;
            }
        }

        /**
         * Given a URL string, return an object { url, headers } with
         * proper headers (Authorization, Map-Agent, etc.), then return it
         * so that fetch() can be called with those headers.
         */
        signRequest(request) {
            request.url = request.url.replace('{azMapsDomain}', this.options.azMapsDomain);
            request.url = request.url.replace('{protocol}', this.options.hasOwnProperty('useSSL') && this.options.useSSL == false ? 'http': 'https');

            const h = Constants;
            const headers = request.headers || {};
            headers[h.SESSION_ID] = AuthenticationManager.sessionId;
            headers[h.MS_AM_REQUEST_ORIGIN] = h.MS_AM_REQUEST_ORIGIN_VALUE;
            headers[h.MAP_AGENT] = `MapControl/${h.SDK_VERSION} (${h.TARGET_SDK})`;

            const token = this.getToken();
            switch (this.options.authType) {
                case 'aad':
                case 'anonymous':
                    headers[h.X_MS_CLIENT_ID] = this.options.clientId;
                    headers[h.AUTHORIZATION] = `${h.AUTHORIZATION_SCHEME} ${token}`;
                    break;
                case 'subscriptionKey':
                    if ('url' in request) {
                        const prefix = request.url.includes('?') ? '&' : '?';
                        request.url += `${prefix}subscription-key=${token}`;
                    } else {
                        throw new Error('No URL specified in request.');
                    }
                    break;
                default:
                    throw new Error('Invalid authentication type.');
            }

            request.headers = headers;
            return request;
        }

        /**
         * Returns a Promise that resolves to a Response after signing:
         * fetch(signedRequest.url, { method:'GET', headers:… }).
         */
        getRequest(url) {
            const req = this.signRequest({ url, headers: {} });
            return fetch(req.url, {
                method: 'GET',
                mode: 'cors',
                headers: new Headers(req.headers)
            });
        }
    }

    // ------------------------------
    // AZURE MAPS TILE GRID
    // ------------------------------
    /**
     * AzureMapsTileGrid extends ol.tilegrid.TileGrid with fixed extent/resolutions
     * and a mutable maxZoom.
     */
    class AzureMapsTileGrid extends ol.tilegrid.TileGrid {
        constructor() {
            super({
                extent: [-20026376.39, -20048966.10, 20026376.39, 20048966.10],
                minZoom: 0,
                origin: [-20037508.342789244, 20037508.342789244],
                resolutions: [
                    156543.03392804097,
                    78271.51696402048,
                    39135.75848201024,
                    19567.87924100512,
                    9783.93962050256,
                    4891.96981025128,
                    2445.98490512564,
                    1222.99245256282,
                    611.49622628141,
                    305.748113140705,
                    152.8740565703525,
                    76.43702828517625,
                    38.21851414258813,
                    19.109257071294063,
                    9.554628535647032,
                    4.777314267823516,
                    2.388657133911758,
                    1.194328566955879,
                    0.5971642834779395,
                    0.29858214173896974,
                    0.14929107086948487,
                    0.07464553543474244,
                    0.03732276771737122
                ],
                tileSize: 256
            });
            this.maxZoom = 22;
        }

        setMaxZoom(z) {
            this.maxZoom = z;
        }

        getMaxZoom() {
            return this.maxZoom;
        }
    }

    // ------------------------------
    // AZURE MAPS SOURCE
    // ------------------------------
    /**
     * ol.source.AzureMaps
     *
     * A tile source for Azure Maps (Render V2, Traffic, Weather, Imagery, etc.).
     * Automatically signs each tile request using subscriptionKey, AAD token, or anonymous token.
     *
     * options = {
     *   tilesetId: string,           // e.g. 'microsoft.base.road'
     *   language: string,            // e.g. 'en-US'
     *   view: string,                // e.g. 'Auto'
     *   timeStamp?: Date|string,     // e.g. '2023-01-01T00:00:00'
     *   trafficFlowThickness?: number,// e.g. 5
     *   authOptions: {
     *     authType: 'subscriptionKey'|'aad'|'anonymous',
     *     subscriptionKey?: string,  // if authType==='subscriptionKey'
     *     getToken?: function,       // if authType==='anonymous'
     *     clientId?: string,         // if authType==='aad' or 'anonymous'
     *     aadTenant?: string,        // if authType==='aad'
     *     aadInstance?: string,      // if authType==='aad'
     *     azMapsDomain?: string      // e.g. 'atlas.microsoft.com'
     *   }
     * }
     */
    class AzureMaps extends ol.source.XYZ {
        constructor(options = {}) {
            // 1) Merge defaults + user options
            const defaultOpts = {
                language: 'en-US',
                view: 'Auto',
                trafficFlowThickness: 5
            };
            const opt = {};
            Object.assign(opt, defaultOpts, options);
           
            // 2) Build the “Render” base URL with the new version:
            //    (We do NOT special-case “/traffic/flow/tile” or “/traffic/incident/tile” anymore.)
            const baseUrl =
                `{protocol}://{azMapsDomain}/map/tile` +
                `?api-version=2024-04-01` +
                `&tilesetId={tilesetId}` +
                `&zoom={z}` +
                `&x={x}` +
                `&y={y}` +
                `&tileSize={tileSize}` +
                `&language={language}` +
                `&view={view}`;

            // 3) Prepare AuthenticationManager as before (subscriptionKey, AAD, or anonymous):
            const au = opt.authOptions || {};
            if (!au.azMapsDomain) {
                au.azMapsDomain = 'atlas.microsoft.com';
            }
            const authMgr = AuthenticationManager.getInstance(au);

            // 4) Create your TileGrid
            const tileGrid = new AzureMapsTileGrid();

            // 5) Call super() WITHOUT referencing `this`
            super({
                projection: 'EPSG:3857',
                url: '', // we’ll set the real URL after auth init
                tileGrid: tileGrid,
                attributions: null,
                tileLoadFunction: null
            });

            // ─── Now you can safely set instance fields ───
            this._options = opt;      // contains tilesetId, language, view, etc.
            this._baseUrl = baseUrl;  // the “/map/tile?…” template
            this._authManager = authMgr;

            // 6) Attach the real attributions & tileLoadFunction:
            this.setAttributions(() => this._getAttribution());
            this.setTileLoadFunction((tile, src) => this._tileLoadFunction(tile, src));

            // 7) Once auth is ready, build & refresh the final tile URL:
            //    (This will replace {tilesetId}, {language}, {view}, etc. and then append ?subscription-key=…)
            if (!this._authManager.isInitialized()) {
                this._authManager.initialize().then(() => {
                    this._refreshTileset();
                });
            } else {
                this._refreshTileset();
            }
        }

        // (… other methods remain unchanged: _getAttribution, _formatUrlTemplate, _refreshTileset, _tileLoadFunction …)
  


        // ─────────────────────────────────────────────────────────────────────────────
        // Public getters/setters
        // ─────────────────────────────────────────────────────────────────────────────
        getLanguage() {
            return this._options.language;
        }
        setLanguage(lang) {
            this._options.language = lang;
            this._refreshTileset();
        }

        getView() {
            return this._options.view;
        }
        setView(view) {
            this._options.view = view;
            this._refreshTileset();
        }

        getTilesetId() {
            return this._options.tilesetId;
        }
        setTilesetId(tilesetId) {
            this._options.tilesetId = tilesetId;

            // rebuild base URL if it’s traffic
            this._baseUrl = `{protocol}://{azMapsDomain}/map/tile?api-version=${Constants.RENDERV2_VERSION}&tilesetId={tilesetId}&zoom={z}&x={x}&y={y}&tileSize={tileSize}&language={language}&view={view}`;
            //if (tilesetId.startsWith('microsoft.traffic.flow')) {
            //    this._baseUrl = 'https://{azMapsDomain}/traffic/flow/tile/png?api-version=1.0&style={style}&zoom={z}&x={x}&y={y}';
            //} else if (tilesetId.startsWith('microsoft.traffic.incident')) {
            //    this._baseUrl = 'https://{azMapsDomain}/traffic/incident/tile/png?api-version=1.0&style={style}&zoom={z}&x={x}&y={y}';
            //}

            // adjust maxZoom for certain tilesets:
            const tg = this.getTileGrid();
            let maxZoom = 22;
            if (tilesetId.startsWith('microsoft.weather.')) {
                maxZoom = 15;
            } else if (tilesetId === 'microsoft.imagery') {
                maxZoom = 19;
            }
            tg.setMaxZoom(maxZoom);

            this._refreshTileset();
        }

        getTimeStamp() {
            return this._options.timeStamp;
        }
        setTimeStamp(timeStamp) {
            this._options.timeStamp = timeStamp;
            this._refreshTileset();
        }

        getTrafficFlowThickness() {
            return this._options.trafficFlowThickness;
        }
        setTrafficFlowThickness(t) {
            this._options.trafficFlowThickness = t;
            this._refreshTileset();
        }

        // ─────────────────────────────────────────────────────────────────────────────
        // Private methods
        // ─────────────────────────────────────────────────────────────────────────────
        _getAttribution() {
            const ts = this._options.tilesetId;
            const year = `© ${new Date().getFullYear()}`;
            let partner = null;
            if (ts) {
                if (ts.startsWith('microsoft.base.') || ts.startsWith('microsoft.traffic.')) {
                    partner = 'TomTom';
                } else if (ts.startsWith('microsoft.weather.')) {
                    partner = 'AccuWeather';
                } else if (ts === 'microsoft.imagery') {
                    partner = 'Airbus';
                }
                if (partner) {
                    return [`${year} ${partner}`, `${year} Microsoft`];
                }
                return `${year} Microsoft`;
            }
            return null;
        }

        _formatUrlTemplate() {
            let url = this._baseUrl
                .replace('{tileSize}', '256')
                .replace('{language}', this._options.language)
                .replace('{view}', this._options.view)
                .replace('{tilesetId}', this._options.tilesetId);

            if (this._options.tilesetId.startsWith('microsoft.traffic.')) {
                const style = this._options.tilesetId
                    .replace('microsoft.traffic.incident.', '')
                    .replace('microsoft.traffic.flow.', '');
                url = url.replace('{style}', style);

                if (this._options.tilesetId.includes('flow')) {
                    url += `&thickness=${this._options.trafficFlowThickness}`;
                }
            }

            if (this._options.timeStamp) {
                let ts = this._options.timeStamp;
                if (ts instanceof Date) {
                    ts = ts.toISOString().slice(0, 19); // remove trailing Z
                }
                url = url.replace('{timeStamp}', ts);
            }

            return url;
        }

        _refreshTileset() {
            const finalUrl = this._formatUrlTemplate();
            // In class syntax, just call `this.setUrl` and `this.refresh`
            this.setUrl(finalUrl);
            this.refresh();
        }

        _tileLoadFunction(imageTile, src) {
            this._authManager
                .getRequest(src)
                .then(response => response.blob())
                .then(blob => {
                    const reader = new FileReader();
                    reader.onload = () => {
                        const imgEl = /** @type {HTMLImageElement} */ (imageTile.getImage());
                        imgEl.src = reader.result;
                    };
                    reader.readAsDataURL(blob);
                })
                .catch(err => {
                    const imgEl = /** @type {HTMLImageElement} */ (imageTile.getImage());
                    imgEl.src = ''; // leave blank on error
                    console.warn('AzureMaps tile failed to load:', err);
                });
        }
    }

    // Expose as ol.source.AzureMaps
    ol.source.AzureMaps = AzureMaps;
})(window.ol);
