/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

let g_version = "0.0.0-0";//"localhost:8000" for develop version
const pathnameParts = self.location.pathname.split('/');
if (pathnameParts.length > 1 && pathnameParts[pathnameParts.length - 2]) {
	g_version = pathnameParts[pathnameParts.length - 2];
}
const g_cacheNamePrefix = 'document_editor_static_';
const g_cacheName = g_cacheNamePrefix + g_version;
const g_cacheablePrefixes = [
	"web-apps/",
	"sdkjs/",
	"fonts/",
	"sdkjs-plugins/",
	"dictionaries/"
];
const isDesktopEditor = navigator.userAgent.indexOf("AscDesktopEditor") !== -1;
let g_storageInfoCache = null;
let g_storageInfoCacheTime = 0;
const STORAGE_INFO_CACHE_DURATION = 30000; // 30 seconds

/**
 * Check if a response is safe to cache
 * @param {Request} request
 * @param {Response} response
 * @returns {boolean} true if response can be safely cached
 */
function safeToCache(request, response) {
	return request.method === 'GET'                                 // only GET requests
		&& response
		&& response.ok                                              // status 200-299. todo 0 or 1223?
		&& (response.type === 'basic' || response.type === 'cors')  // same-origin or CORS
		&& !response.redirected;                  	                // no 30x redirect chain
}

/**
 * Get storage information (size limits and health) with single API call
 * @returns {Promise<{maxEntrySize: number, isHealthy: boolean}>} Storage info
 */
function getStorageInfo() {
	const now = Date.now();
	if (g_storageInfoCache !== null && (now - g_storageInfoCacheTime) < STORAGE_INFO_CACHE_DURATION) {
		return Promise.resolve(g_storageInfoCache);
	}
	
	if (!navigator.storage || !navigator.storage.estimate) {
		// Fallback values if API not available
		g_storageInfoCache = {
			maxEntrySize: 50 * 1024 * 1024,
			isHealthy: true
		};
		g_storageInfoCacheTime = now;
		return Promise.resolve(g_storageInfoCache);
	}
	
	return navigator.storage.estimate()
		.then(function(estimate) {
			// Validate estimate fields; fall back if missing or invalid
			if (!estimate || typeof estimate.quota !== 'number' || !isFinite(estimate.quota) || estimate.quota <= 0 ||
				typeof estimate.usage !== 'number' || !isFinite(estimate.usage)) {
				g_storageInfoCache = {
					maxEntrySize: 50 * 1024 * 1024,
					isHealthy: true
				};
				g_storageInfoCacheTime = Date.now();
				return g_storageInfoCache;
			}
			// Calculate max entry size: cache ≈ 10% of quota, cap entry at 1/8th
			const cacheSize = Math.min(estimate.quota * 0.10, 1024 * 1024 * 1024); // 1 GiB max
			const maxEntrySize = cacheSize / 8; // Per-entry cap is 1/8th of cache size
			
			// Calculate storage health: back off when disk is 80% full
			const usageRatio = estimate.usage / estimate.quota;
			const isHealthy = usageRatio < 0.8;
			
			g_storageInfoCache = { maxEntrySize: maxEntrySize, isHealthy: isHealthy };
			g_storageInfoCacheTime = Date.now();
			return g_storageInfoCache;
		})
		.catch(function(error) {
			// Fallback values on error
			g_storageInfoCache = {
				maxEntrySize: 50 * 1024 * 1024,
				isHealthy: true
			};
			g_storageInfoCacheTime = Date.now();
			return g_storageInfoCache;
		});
}

/**
 * Check if response size is within cacheable limits
 * @param {Response} response
 * @returns {Promise<boolean>} true if response is not too large
 */
function isReasonableSize(response) {
	const size = Number(response.headers.get('content-length')) || 0;
	if (size === 0) {
		return Promise.resolve(true); // No size header, assume reasonable
	}
	
	return getStorageInfo()
		.then(function(info) {
			return size < info.maxEntrySize;
		});
}

/**
 * Check if storage quota is healthy for caching
 * @returns {Promise<boolean>} true if storage is healthy
 */
function storageHealthy() {
	return getStorageInfo()
		.then(function(info) {
			return info.isHealthy;
		});
}

/**
 * Put response in cache with retry logic for transient errors
 * @param {Request} request
 * @param {Response} response - Response to cache; this function clones per attempt to preserve the original for retries
 * @param {number} attempt - Current attempt number (for retry logic)
 * @returns {Promise} Promise that resolves when caching completes or fails
 */
function putInCache(request, response, attempt) {
    if (typeof attempt === 'undefined') attempt = 0;
    return caches.open(g_cacheName)
        .then(function(cache) {
            // Clone at the moment of caching so the provided response remains pristine for retries
            return cache.put(request, response.clone());
        })
        .catch(function(err) {
            // Transient quota/disk hiccup? Retry up to 2x with exponential back-off
            if (attempt < 2) {
                return new Promise(function(resolve) {
                    setTimeout(resolve, 250 * Math.pow(2, attempt)); // 250ms, 500ms
                })
                .then(function() {
                    // Reuse the original unconsumed response; a fresh clone will be created inside cache.put
                    return putInCache(request, response, attempt + 1);
                });
            } else {
                const size = response.headers ? response.headers.get('content-length') : 'unknown';
                console.error('putInCache failed after max retries:', {
                    url: request.url,
                    method: request.method,
                    responseSize: size,
                    responseType: response.type,
                    cacheName: g_cacheName,
                    error: err && (err.message || err)
                });
            }
        });
}

function cacheFirst(event) {
	const request = event.request;
	
	return caches.match(request, { cacheName: g_cacheName })
		.then(function(cached) {
			return cached || fetch(request).then(function(networkResp) {
			// Clone immediately to avoid "body already used" errors
			const responseForCache = networkResp.clone();
			
			// Fire-and-forget caching **after** responding to the page
			if (safeToCache(request, networkResp)) {
				event.waitUntil(
					getStorageInfo()
					.then(function(info) {
						const size = Number(networkResp.headers.get('content-length')) || 0;
						const sizeOk = size === 0 || size < info.maxEntrySize;
						
						if (info.isHealthy && sizeOk) {
							return putInCache(request, responseForCache);
						}
					})
				);
			}
			return networkResp;
			});
		});
}
function activateWorker(event) {
	return self.clients.claim()
		.then(function(){
			//remove stale caches
			return caches.keys();
		})
		.then(function (keys) {
			const deletePromises = keys
				.filter(function(cache) {
					return cache.includes(g_cacheNamePrefix) && !cache.includes(g_cacheName);
				})
				.map(function(cache) {
					return caches.delete(cache);
				});
			return Promise.all(deletePromises);
		}).catch(function (err) {
			console.error('activateWorker failed with ' + err);
		});
}
/**
 * Filter function for cacheable paths
 * @param {string} url
 * @returns {boolean}
 */
function matchesCacheablePath(url) {
	const g_versionNeedle = "/" + g_version + "/";
    const versionIndex = url.indexOf(g_versionNeedle);
    if (versionIndex === -1) return false;

    // Position just after "/<version>/"
    const i = versionIndex + g_versionNeedle.length;

	for (let k = 0; k < g_cacheablePrefixes.length; k++) {
		//startsWith not supported in ie11 but at first service worker not supported
		if (url.startsWith(g_cacheablePrefixes[k], i)) {
            return true;
        }
    }
    return false;
}

self.addEventListener('install', (event) => {
	event.waitUntil(self.skipWaiting());
});

self.addEventListener('activate', (event) => {
	event.waitUntil(activateWorker());
});

self.addEventListener('fetch', (event) => {
	const request = event.request;
	const url = request.url;
	
	// Fast path: check method and URL pattern in one go
	if (request.method !== "GET" || !matchesCacheablePath(url)) {
		return;
	}

	// Desktop editor exclusion
	if (isDesktopEditor && url.indexOf("/sdkjs/common/AllFonts.js") !== -1) {
		return;
	}

	event.respondWith(cacheFirst(event));
});
