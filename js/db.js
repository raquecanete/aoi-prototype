(function (global) {
  'use strict';

  var DB_NAME = 'aoi-workbook-library';
  var DB_VERSION = 1;
  var STORE_KEYS = {
    files: 'id',
    versions: 'id',
    drafts: 'fileId',
    settings: 'key'
  };
  var storeNames = Object.keys(STORE_KEYS);
  var cache = {};
  var database = null;
  var currentMode = 'indexeddb';
  var warningShown = false;
  var readySettled = false;
  var snapshotPromise = null;
  var resolveReady;

  storeNames.forEach(function (name) {
    cache[name] = new Map();
  });

  var ready = new Promise(function (resolve) {
    resolveReady = resolve;
  });

  function clone(value) {
    if (value === undefined || value === null) return value;
    if (typeof global.structuredClone === 'function') {
      try {
        return global.structuredClone(value);
      } catch (error) {
        // Fall through for data types unsupported by structuredClone.
      }
    }
    try {
      return JSON.parse(JSON.stringify(value));
    } catch (error) {
      return value;
    }
  }

  function resolveReadyOnce() {
    if (readySettled) return;
    readySettled = true;
    resolveReady(api);
  }

  function showStorageWarning(error) {
    if (warningShown) return;
    warningShown = true;

    var message = 'Browser storage is unavailable. Changes are being kept in memory and may not persist after you close this page.';
    var warning = global.document && global.document.getElementById('storageWarning');
    if (warning) {
      warning.textContent = message;
      warning.hidden = false;
      if (warning.style) warning.style.display = '';
      if (warning.setAttribute) warning.setAttribute('role', 'alert');
    }

    if (typeof global.dispatchEvent === 'function' && typeof global.CustomEvent === 'function') {
      global.dispatchEvent(new global.CustomEvent('storage-warning', {
        detail: { message: message, error: error || null }
      }));
    }
  }

  function activateMemoryMode(error) {
    if (currentMode !== 'memory') {
      currentMode = 'memory';
      showStorageWarning(error);
      if (database) snapshotPromise = snapshotDatabase(database);
    }
    resolveReadyOnce();
    return snapshotPromise || Promise.resolve();
  }

  function snapshotDatabase(db) {
    return new Promise(function (resolve) {
      var transaction;
      try {
        transaction = db.transaction(storeNames, 'readonly');
        storeNames.forEach(function (name) {
          var request = transaction.objectStore(name).getAll();
          request.onsuccess = function () {
            request.result.forEach(function (value) {
              var key = value && value[STORE_KEYS[name]];
              if (key !== undefined) cache[name].set(key, clone(value));
            });
          };
        });
        transaction.oncomplete = function () {
          db.close();
          if (database === db) database = null;
          resolve();
        };
        transaction.onerror = transaction.onabort = function () {
          db.close();
          if (database === db) database = null;
          resolve();
        };
      } catch (snapshotError) {
        db.close();
        if (database === db) database = null;
        resolve();
      }
    });
  }

  function makeTransactionError(transaction, fallbackMessage) {
    var cause = transaction && transaction.error;
    var error = new Error((cause && cause.message) || fallbackMessage);
    error.name = (cause && cause.name) || 'TransactionError';
    if (cause) error.cause = cause;
    return error;
  }

  function isRecoverable(error) {
    return !!error && (
      error.name === 'QuotaExceededError' ||
      error.name === 'SecurityError' ||
      error.name === 'InvalidStateError' ||
      error.name === 'NotSupportedError'
    );
  }

  function openDatabase() {
    var request;
    try {
      if (!global.indexedDB) throw new Error('IndexedDB is unavailable');
      request = global.indexedDB.open(DB_NAME, DB_VERSION);
    } catch (error) {
      activateMemoryMode(error);
      return;
    }

    request.onupgradeneeded = function (event) {
      var db = event.target.result;
      if (!db.objectStoreNames.contains('files')) {
        db.createObjectStore('files', { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains('versions')) {
        var versions = db.createObjectStore('versions', { keyPath: 'id' });
        versions.createIndex('fileId', 'fileId', { unique: false });
      } else {
        var versionStore = event.target.transaction.objectStore('versions');
        if (!versionStore.indexNames.contains('fileId')) {
          versionStore.createIndex('fileId', 'fileId', { unique: false });
        }
      }
      if (!db.objectStoreNames.contains('drafts')) {
        db.createObjectStore('drafts', { keyPath: 'fileId' });
      }
      if (!db.objectStoreNames.contains('settings')) {
        db.createObjectStore('settings', { keyPath: 'key' });
      }
    };

    request.onblocked = function () {
      activateMemoryMode(new Error('Opening IndexedDB was blocked'));
    };

    request.onerror = function () {
      activateMemoryMode(request.error || new Error('Unable to open IndexedDB'));
    };

    request.onsuccess = function () {
      if (currentMode === 'memory') {
        request.result.close();
        return;
      }
      database = request.result;
      database.onversionchange = function () {
        activateMemoryMode(new Error('IndexedDB version changed'));
      };
      hydrateCache();
    };
  }

  function hydrateCache() {
    var transaction;
    try {
      transaction = database.transaction(['files', 'settings'], 'readonly');
    } catch (error) {
      activateMemoryMode(error);
      return;
    }

    var failed = false;
    ['files', 'settings'].forEach(function (name) {
      var request = transaction.objectStore(name).getAll();
      request.onsuccess = function () {
        request.result.forEach(function (value) {
          var key = value && value[STORE_KEYS[name]];
          if (key !== undefined) cache[name].set(key, clone(value));
        });
      };
      request.onerror = function () {
        failed = true;
      };
    });

    transaction.oncomplete = function () {
      if (!failed) resolveReadyOnce();
    };
    transaction.onerror = transaction.onabort = function () {
      activateMemoryMode(makeTransactionError(transaction, 'Unable to load stored data'));
    };
  }

  function assertStore(store) {
    if (!Object.prototype.hasOwnProperty.call(STORE_KEYS, store)) {
      throw new Error('Unknown store: ' + store);
    }
  }

  function transact(store, mode, action, fallbackAction) {
    return ready.then(function () {
      assertStore(store);
      if (currentMode === 'memory') {
        return Promise.resolve(snapshotPromise).then(fallbackAction);
      }

      return new Promise(function (resolve, reject) {
        var transaction;
        var result;
        var actionError = null;
        try {
          transaction = database.transaction([store], mode);
          result = action(transaction.objectStore(store));
          if (result && typeof result.then === 'function') {
            result.then(function (value) {
              result = value;
            }, function (error) {
              actionError = error;
            });
          }
        } catch (error) {
          if (isRecoverable(error)) {
            activateMemoryMode(error).then(function () {
              try {
                resolve(fallbackAction());
              } catch (fallbackError) {
                reject(fallbackError);
              }
            });
          } else {
            reject(error);
          }
          return;
        }

        transaction.oncomplete = function () {
          if (actionError) reject(actionError);
          else resolve(result);
        };
        transaction.onabort = function () {
          var error = makeTransactionError(transaction, 'IndexedDB transaction failed');
          if (isRecoverable(error) || (transaction.error && isRecoverable(transaction.error))) {
            activateMemoryMode(error).then(function () {
              try {
                resolve(fallbackAction());
              } catch (fallbackError) {
                reject(fallbackError);
              }
            });
          } else {
            reject(actionError || error);
          }
        };
      });
    });
  }

  var api = {
    ready: ready,

    get: function (store, key) {
      return transact(store, 'readonly', function (objectStore) {
        return new Promise(function (resolve, reject) {
          var request = objectStore.get(key);
          request.onsuccess = function () {
            resolve(clone(request.result));
          };
          request.onerror = function () {
            reject(request.error || new Error('Unable to read from ' + store));
          };
        });
      }, function () {
        return clone(cache[store].get(key));
      }).then(function (value) {
        if (value === undefined) cache[store].delete(key);
        else cache[store].set(key, clone(value));
        return value;
      });
    },

    getAll: function (store) {
      return transact(store, 'readonly', function (objectStore) {
        return new Promise(function (resolve, reject) {
          var request = objectStore.getAll();
          request.onsuccess = function () {
            resolve(request.result.map(clone));
          };
          request.onerror = function () {
            reject(request.error || new Error('Unable to read from ' + store));
          };
        });
      }, function () {
        return Array.from(cache[store].values()).map(clone);
      }).then(function (values) {
        cache[store].clear();
        values.forEach(function (value) {
          var key = value && value[STORE_KEYS[store]];
          if (key !== undefined) cache[store].set(key, clone(value));
        });
        return values;
      });
    },

    put: function (store, value) {
      return ready.then(function () {
        assertStore(store);
        var keyPath = STORE_KEYS[store];
        if (!value || value[keyPath] === undefined || value[keyPath] === null) {
          throw new Error('A value for ' + store + ' must include ' + keyPath);
        }
        var savedValue = clone(value);
        var key = savedValue[keyPath];
        if (currentMode === 'memory') {
          cache[store].set(key, savedValue);
          return clone(savedValue);
        }
        return transact(store, 'readwrite', function (objectStore) {
          var request = objectStore.put(savedValue);
          return new Promise(function (resolve, reject) {
            request.onsuccess = function () {
              resolve(clone(savedValue));
            };
            request.onerror = function () {
              reject(request.error || new Error('Unable to save to ' + store));
            };
          });
        }, function () {
          cache[store].set(key, savedValue);
          return clone(savedValue);
        }).then(function (saved) {
          cache[store].set(key, clone(saved));
          return saved;
        });
      });
    },

    delete: function (store, key) {
      return transact(store, 'readwrite', function (objectStore) {
        var request = objectStore.delete(key);
        return new Promise(function (resolve, reject) {
          request.onsuccess = function () {
            resolve();
          };
          request.onerror = function () {
            reject(request.error || new Error('Unable to delete from ' + store));
          };
        });
      }, function () {
        cache[store].delete(key);
      }).then(function (result) {
        cache[store].delete(key);
        return result;
      });
    },

    clear: function (store) {
      return transact(store, 'readwrite', function (objectStore) {
        var request = objectStore.clear();
        return new Promise(function (resolve, reject) {
          request.onsuccess = function () {
            resolve();
          };
          request.onerror = function () {
            reject(request.error || new Error('Unable to clear ' + store));
          };
        });
      }, function () {
        cache[store].clear();
      }).then(function (result) {
        cache[store].clear();
        return result;
      });
    },

    getVersionsForFile: function (fileId) {
      return ready.then(function () {
        if (currentMode === 'memory') {
          return Promise.resolve(snapshotPromise).then(function () {
            return Array.from(cache.versions.values())
              .filter(function (version) { return version.fileId === fileId; })
              .map(clone);
          });
        }
        return new Promise(function (resolve, reject) {
          var transaction;
          try {
            transaction = database.transaction(['versions'], 'readonly');
            var request = transaction.objectStore('versions').index('fileId').getAll(fileId);
            request.onsuccess = function () {
              var versions = request.result.map(clone);
              versions.forEach(function (version) { cache.versions.set(version.id, clone(version)); });
              resolve(versions);
            };
            request.onerror = function () {
              reject(request.error || new Error('Unable to load file versions'));
            };
          } catch (error) {
            if (isRecoverable(error)) {
              activateMemoryMode(error).then(function () {
                resolve(Array.from(cache.versions.values())
                  .filter(function (version) { return version.fileId === fileId; })
                  .map(clone));
              });
            } else reject(error);
          }
        });
      });
    },

    estimate: function () {
      return ready.then(function () {
        var storage = global.navigator && global.navigator.storage;
        if (!storage || typeof storage.estimate !== 'function') {
          return { usage: null, quota: null };
        }
        try {
          return Promise.resolve(storage.estimate()).catch(function () {
            return { usage: null, quota: null };
          });
        } catch (error) {
          return { usage: null, quota: null };
        }
      });
    }
  };

  Object.defineProperty(api, 'mode', {
    enumerable: true,
    get: function () {
      return currentMode;
    }
  });

  global.AOIDB = api;
  openDatabase();
}(window));
