/**
 * @NApiVersion 2.1
 * @NScriptType MapReduceScript
 */
define(['N/file', 'N/search', 'N/https', 'N/log', 'N/runtime', 'N/encode'], function (file, search, https, log, runtime, encode) {

    function getInputData() {
        var scriptObj = runtime.getCurrentScript();
        var folderParam = (scriptObj.getParameter({ name: 'custscript_file_cabinate_folder' }) || '').trim();
        var folderMappingParam = (scriptObj.getParameter({ name: 'custscript_sp_folder_mapping' }) || '').trim();

        // Parse comma-separated folder IDs
        var rootFolderIds = folderParam.split(',')
            .map(function (id) { return id.trim(); })
            .filter(function (id) { return id.length > 0; })
            .map(function (id) { return String(id); });

        log.debug('Input Data', 'Searching folder IDs: ' + (JSON.stringify(rootFolderIds)));

        if (rootFolderIds.length === 0) {
            log.audit('Input Data', 'No folder IDs specified.');
            return [];
        }

        // Parse mapping parameter if provided
        var folderMapping = null;
        if (folderMappingParam) {
            try {
                folderMapping = JSON.parse(folderMappingParam);
                log.debug('Folder Mapping Parsed', 'Total mapped file IDs: ' + (Object.keys(folderMapping).length));
            } catch (e) {
                log.error('Invalid Folder Mapping JSON', e.message);
            }
        }

        // Get recursive hierarchy
        var hierarchy = getFolderHierarchy(rootFolderIds);

        log.debug('Hierarchy details', 'Total folders to search: ' + (hierarchy.allFolderIds.length));

        // Search files in all these folder IDs
        var fileSearch = search.create({
            type: 'file',
            filters: [
                ['folder', 'anyof', hierarchy.allFolderIds]
            ],
            columns: ['internalid', 'name', 'folder']
        });

        var fileResults = [];

        // Paginate and fetch all files
        var pageData = fileSearch.runPaged({ pageSize: 1000 });
        pageData.pageRanges.forEach(function (pageRange) {
            var page = pageData.fetch({ index: pageRange.index });
            page.data.forEach(function (result) {
                var fileId = result.id;
                var fileName = result.getValue('name');
                var fileFolder = result.getValue('folder') ? String(result.getValue('folder')) : '';

                var relativePath = '';
                if (folderMapping) {
                    if (folderMapping[fileId]) {
                        relativePath = folderMapping[fileId];
                        fileResults.push({
                            id: fileId,
                            name: fileName,
                            folder: fileFolder,
                            relativePath: relativePath
                        });
                    } else {
                        log.debug('Skip File', 'File ID ' + (fileId) + ' (' + (fileName) + ') not found in folder mapping.');
                    }
                } else {
                    // Fallback to relative path by folder structure if no mapping provided
                    relativePath = getRelativeFolderPath(fileFolder, rootFolderIds, hierarchy.parentMap, hierarchy.nameMap);
                    fileResults.push({
                        id: fileId,
                        name: fileName,
                        folder: fileFolder,
                        relativePath: relativePath
                    });
                }
            });
        });

        log.audit('Input Data Complete', 'Total files found for processing: ' + (fileResults.length));
        return fileResults;
    }

    function map(context) {
        var scriptObj = runtime.getCurrentScript();
        var siteId = (scriptObj.getParameter({ name: 'custscript_site_id' }) || '').trim();
        var sharepointFolderPath = (scriptObj.getParameter({ name: 'custscript_sharepoint_folder_path' }) || '').trim();

        var result = JSON.parse(context.value);
        var fileId = result.id;
        var fileName = result.name;
        var relativePath = result.relativePath;

        // Load file from NetSuite
        var nsFile = file.load({ id: fileId });

        try {
            var accessToken = getAccessToken(scriptObj);
            if (!accessToken) throw new Error('Could not retrieve SharePoint access token.');
            accessToken = accessToken.trim();

            // Sanitize siteId: ensure it doesn't start with a slash
            var sanitizedSiteId = siteId.replace(/^\/+/, '');

            // Sanitize and encode the path
            var fileNameEncoded = encodeURIComponent(fileName);

            // Build the final folder path for SharePoint
            var finalFolderPath = sharepointFolderPath;
            if (relativePath) {
                if (finalFolderPath) {
                    finalFolderPath = finalFolderPath.replace(/\/+$/, '') + '/' + relativePath;
                } else {
                    finalFolderPath = relativePath;
                }
            }

            // Remove leading and trailing slashes from the folder path
            finalFolderPath = finalFolderPath.replace(/^\/+|\/+$/g, '');

            var uploadUrl;
            var folderEncoded = '';
            if (finalFolderPath) {
                // Encode each segment of the folder path to preserve slashes but escape special characters
                folderEncoded = finalFolderPath.split('/')
                    .map(function (segment) { return encodeURIComponent(segment); })
                    .join('/');
            }

            // Check if the SharePoint folder exists before uploading
            var folderExists = false;
            if (finalFolderPath) {
                var folderCheckUrl = 'https://graph.microsoft.com/v1.0/sites/' + (sanitizedSiteId) + '/drive/root:/' + (folderEncoded);

                var checkResponse = https.get({
                    url: folderCheckUrl,
                    headers: {
                        'Authorization': 'Bearer ' + accessToken,
                        'Accept': 'application/json'
                    }
                });

                log.debug('Folder Check Response', { folder: finalFolderPath, code: checkResponse.code });

                if (checkResponse.code === 200) {
                    // Confirm it is actually a folder
                    try {
                        var body = JSON.parse(checkResponse.body);
                        if (body.folder) {
                            folderExists = true;
                        } else {
                            log.error('Folder Check', 'Path exists but is not a folder: ' + (finalFolderPath));
                        }
                    } catch (e) {
                        folderExists = true;
                    }
                }
            } else {
                // Root folder always exists
                folderExists = true;
            }

            if (!folderExists) {
                log.audit('Folder Not Found', 'Folder \'' + (finalFolderPath) + '\' was not found in SharePoint. Skipping upload for file: ' + (fileName));
                context.write({
                    key: fileId,
                    value: JSON.stringify({ status: 'skipped', name: fileName, reason: 'Folder \'' + (finalFolderPath) + '\' not found in SharePoint.' })
                });
                return;
            }

            var fileSize = nsFile.size;
            var fileType = nsFile.fileType;
            var textTypes = ['PLAINTEXT', 'CSV', 'JSON', 'JAVASCRIPT', 'XMLDOC', 'HTMLDOC', 'STYLESHEET', 'MESSAGES', 'RTF', 'SMS'];
            var isText = textTypes.indexOf(fileType) !== -1;

            // Files larger than 4MB must use an upload session
            var useUploadSession = fileSize > 4 * 1024 * 1024;

            if (useUploadSession) {
                if (!isText) {
                    throw new Error('File size is ' + (fileSize) + ' bytes and exceeds 4MB, but chunked upload is only supported for text/CSV files.');
                }

                log.audit('Starting Upload Session', 'File: ' + (fileName) + ', Size: ' + (fileSize) + ' bytes');

                // 1. Create Upload Session
                var sessionUrl;
                if (folderEncoded) {
                    sessionUrl = 'https://graph.microsoft.com/v1.0/sites/' + (sanitizedSiteId) + '/drive/root:/' + (folderEncoded) + '/' + (fileNameEncoded) + ':/createUploadSession';
                } else {
                    sessionUrl = 'https://graph.microsoft.com/v1.0/sites/' + (sanitizedSiteId) + '/drive/root:/' + (fileNameEncoded) + ':/createUploadSession';
                }

                var sessionResponse = https.post({
                    url: sessionUrl,
                    headers: {
                        'Authorization': 'Bearer ' + accessToken,
                        'Content-Type': 'application/json',
                        'Accept': 'application/json'
                    },
                    body: JSON.stringify({
                        item: {
                            "@microsoft.graph.conflictBehavior": "replace"
                        }
                    })
                });

                // Fallback / retry logic if we encounter 409 Conflict (nameAlreadyExists / upload session lock)
                if (sessionResponse.code === 409) {
                    log.audit('Conflict Encountered', 'File \'' + (fileName) + '\' already exists or upload is locked. Attempting to delete and retry...');

                    var deleteUrl;
                    if (folderEncoded) {
                        deleteUrl = 'https://graph.microsoft.com/v1.0/sites/' + (sanitizedSiteId) + '/drive/root:/' + (folderEncoded) + '/' + (fileNameEncoded);
                    } else {
                        deleteUrl = 'https://graph.microsoft.com/v1.0/sites/' + (sanitizedSiteId) + '/drive/root:/' + (fileNameEncoded);
                    }

                    var deleteResponse = https.delete({
                        url: deleteUrl,
                        headers: {
                            'Authorization': 'Bearer ' + accessToken,
                            'Accept': 'application/json'
                        }
                    });

                    log.debug('Delete File Response', { file: fileName, code: deleteResponse.code });

                    // Retry creating the upload session
                    sessionResponse = https.post({
                        url: sessionUrl,
                        headers: {
                            'Authorization': 'Bearer ' + accessToken,
                            'Content-Type': 'application/json',
                            'Accept': 'application/json'
                        },
                        body: JSON.stringify({
                            item: {
                                "@microsoft.graph.conflictBehavior": "replace"
                            }
                        })
                    });
                }

                if (sessionResponse.code !== 200) {
                    throw new Error('Failed to create upload session: HTTP ' + (sessionResponse.code) + ': ' + (sessionResponse.body));
                }

                var sessionData = JSON.parse(sessionResponse.body);
                var sessionUploadUrl = sessionData.uploadUrl;

                var chunkSize = 327680 * 10; // ~3.12 MB
                var reader = nsFile.getReader();
                var charsPerRead = 100000; // read roughly one upload-chunk worth of chars at a time
                var bytesUploaded = 0;
                var pendingHex = '';

                while (bytesUploaded + (pendingHex.length / 2) < fileSize) {
                    var textChunk = reader.readChars({ number: charsPerRead });
                    if (!textChunk) break; // nothing left to read

                    var hexPart = encode.convert({
                        string: textChunk,
                        inputEncoding: encode.Encoding.UTF_8,
                        outputEncoding: encode.Encoding.HEX
                    });
                    pendingHex += hexPart;

                    // Flush full (or final) chunks to SharePoint
                    while (pendingHex.length / 2 >= chunkSize ||
                          (pendingHex.length > 0 && bytesUploaded + pendingHex.length / 2 >= fileSize)) {

                        var thisChunkBytes = Math.min(chunkSize, pendingHex.length / 2);
                        var hexSlice = pendingHex.substring(0, thisChunkBytes * 2);
                        pendingHex = pendingHex.substring(thisChunkBytes * 2);

                        var chunkData = hexToUint8Array(hexSlice);
                        var chunkStart = bytesUploaded;
                        var chunkEnd = bytesUploaded + thisChunkBytes - 1;
                        var isLastChunk = chunkEnd >= fileSize - 1;

                        var rangeHeader = 'bytes ' + chunkStart + '-' + chunkEnd + '/' + fileSize;

                        log.debug('Uploading Chunk', {
                            file: fileName,
                            start: chunkStart,
                            end: chunkEnd,
                            size: thisChunkBytes,
                            total: fileSize
                        });

                        var chunkResponse = https.put({
                            url: sessionUploadUrl,
                            headers: {
                                'Content-Length': String(thisChunkBytes),
                                'Content-Range': rangeHeader,
                                'Content-Type': 'application/octet-stream'
                            },
                            body: chunkData
                        });

                        if (chunkResponse.code === 200 || chunkResponse.code === 201) {
                            if (isLastChunk) {
                                context.write({ key: fileId, value: JSON.stringify({ status: 'success', name: fileName }) });
                            }
                        } else if (chunkResponse.code !== 202) {
                            throw new Error('Failed to upload chunk range ' + rangeHeader + ': HTTP ' + chunkResponse.code + ': ' + chunkResponse.body);
                        }

                        bytesUploaded += thisChunkBytes;
                    }
                }

                // Post-loop safety check for any remaining bytes
                if (pendingHex.length > 0) {
                    var thisChunkBytes = pendingHex.length / 2;
                    var chunkData = hexToUint8Array(pendingHex);
                    var chunkStart = bytesUploaded;
                    var chunkEnd = bytesUploaded + thisChunkBytes - 1;
                    var rangeHeader = 'bytes ' + chunkStart + '-' + chunkEnd + '/' + fileSize;

                    log.debug('Uploading Final Chunk (Safety)', {
                        file: fileName,
                        start: chunkStart,
                        end: chunkEnd,
                        size: thisChunkBytes,
                        total: fileSize
                    });

                    var chunkResponse = https.put({
                        url: sessionUploadUrl,
                        headers: {
                            'Content-Length': String(thisChunkBytes),
                            'Content-Range': rangeHeader,
                            'Content-Type': 'application/octet-stream'
                        },
                        body: chunkData
                    });

                    if (chunkResponse.code === 200 || chunkResponse.code === 201) {
                        context.write({ key: fileId, value: JSON.stringify({ status: 'success', name: fileName }) });
                    } else {
                        throw new Error('Failed to upload final chunk range ' + rangeHeader + ': HTTP ' + chunkResponse.code + ': ' + chunkResponse.body);
                    }
                    bytesUploaded += thisChunkBytes;
                }
            } else {
                // Direct upload for smaller files
                var fileContent;
                if (isText) {
                    fileContent = nsFile.getContents();
                } else {
                    var base64Content = nsFile.getContents();
                    var hexContent = encode.convert({
                        string: base64Content,
                        inputEncoding: encode.Encoding.BASE_64,
                        outputEncoding: encode.Encoding.HEX
                    });
                    fileContent = hexToUint8Array(hexContent);
                }

                var uploadUrl;
                if (folderEncoded) {
                    uploadUrl = 'https://graph.microsoft.com/v1.0/sites/' + (sanitizedSiteId) + '/drive/root:/' + (folderEncoded) + '/' + (fileNameEncoded) + ':/content';
                } else {
                    uploadUrl = 'https://graph.microsoft.com/v1.0/sites/' + (sanitizedSiteId) + '/drive/root:/' + (fileNameEncoded) + ':/content';
                }

                var response = https.put({
                    url: uploadUrl,
                    body: fileContent,
                    headers: {
                        'Authorization': 'Bearer ' + accessToken,
                        'Content-Type': 'application/octet-stream',
                        'Accept': '*/*'
                    }
                });

                log.debug('Direct Upload Response', { file: fileName, code: response.code });

                if (response.code === 200 || response.code === 201) {
                    context.write({
                        key: fileId,
                        value: JSON.stringify({ status: 'success', name: fileName })
                    });
                } else {
                    context.write({
                        key: fileId,
                        value: JSON.stringify({ status: 'failure', name: fileName, error: 'HTTP ' + (response.code) + ': ' + (response.body) })
                    });
                }
            }
        } catch (err) {
            var errMsg = err.message || err.toString();
            log.error('Process Error', 'File: ' + (fileName) + ' Error: ' + (errMsg));
            context.write({
                key: fileId,
                value: JSON.stringify({ status: 'failure', name: fileName, error: errMsg })
            });
        }
    }

    function summarize(summary) {
        var successCount = 0;
        var failureCount = 0;
        var skipCount = 0;
        var successfulFiles = [];
        var failedFiles = [];
        var skippedFiles = [];

        // Check for any input stage errors
        if (summary.inputSummary.error) {
            log.error('Input Stage Error', summary.inputSummary.error);
        }

        // Check for map stage errors
        summary.mapSummary.errors.iterator().each(function (key, error) {
            log.error('Map Stage System Error for Key: ' + key, error);
            return true;
        });

        // Iterate over outputs written in Map stage
        summary.output.iterator().each(function (key, value) {
            try {
                var result = JSON.parse(value);
                if (result.status === 'success') {
                    successCount++;
                    successfulFiles.push(result.name);
                } else if (result.status === 'failure') {
                    failureCount++;
                    failedFiles.push((result.name) + ' (' + (result.error) + ')');
                } else if (result.status === 'skipped') {
                    skipCount++;
                    skippedFiles.push((result.name) + ' (' + (result.reason) + ')');
                }
            } catch (e) {
                log.error('Summarize Parse Error', 'Key: ' + (key) + ', Value: ' + (value) + ', Error: ' + (e.message));
            }
            return true;
        });

        // Produce a clean, single-audit log summary
        log.audit('Upload Summary',
            'Upload Completed.\n' +
            'Total Files: ' + (successCount + failureCount + skipCount) + '\n' +
            'Success Count: ' + (successCount) + '\n' +
            'Failure Count: ' + (failureCount) + '\n' +
            'Skipped Count: ' + (skipCount) + '\n\n' +
            'Successful Files: [ ' + (successfulFiles.join(', ')) + ' ]\n\n' +
            'Failed Files: [ ' + (failedFiles.join('; ')) + ' ]\n\n' +
            'Skipped Files: [ ' + (skippedFiles.join('; ')) + ' ]'
        );
    }

    /**
     * Helper to get OAuth token using client credentials and secrets
     */
    function getAccessToken(scriptObj) {
        try {
            var tenantId = (scriptObj.getParameter({ name: 'custscript_tenant_id' }) || '').trim();
            if (!tenantId) throw new Error('Tenant ID is missing in script parameters.');

            if (tenantId.toLowerCase() === 'common') {
                log.error('Config Error', 'Client Credentials flow does not support "common" tenant. Please use a specific Tenant ID (GUID).');
            }

            var tokenUrl = 'https://login.microsoftonline.com/' + (tenantId) + '/oauth2/v2.0/token';

            var bodyPayload =
                'client_id={custsecret_sp_client_id}' +
                '&client_secret={custsecret_sp_client_secret}' +
                '&scope=https://graph.microsoft.com/.default' +
                '&grant_type=client_credentials';

            var secureBody = https.createSecureString({ input: bodyPayload });

            var response = https.post({
                url: tokenUrl,
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                    'Accept': 'application/json'
                },
                body: secureBody
            });

            if (response.code !== 200) {
                log.error('Token Retrieval Error', response.body);
                return null;
            }

            return JSON.parse(response.body).access_token;
        } catch (e) {
            log.error('getAccessToken Error', e.message);
            return null;
        }
    }

    var hexLookup = new Uint8Array(256);
    (function() {
        var chars = "0123456789abcdefABCDEF";
        for (var i = 0; i < chars.length; i++) {
            var code = chars.charCodeAt(i);
            var val = i;
            if (val >= 16) val -= 6; // ABCDEF map to 10-15
            hexLookup[code] = val;
        }
    })();

    /**
     * Highly optimized helper to convert Hex string representation to Uint8Array for binary transmission.
     * Uses a pre-allocated lookup table and charCodeAt to bypass the JavaScript statement limit.
     */
    function hexToUint8Array(hex) {
        var len = hex.length;
        var bytes = new Uint8Array(len / 2);
        for (var i = 0; i < len; i += 2) {
            var high = hexLookup[hex.charCodeAt(i)];
            var low = hexLookup[hex.charCodeAt(i + 1)];
            bytes[i / 2] = (high << 4) | low;
        }
        return bytes;
    }

    /**
     * Recursively traverses and compiles NetSuite folder hierarchy down from rootFolderIds
     */
    function getFolderHierarchy(rootFolderIds) {
        var allFolderIds = [...rootFolderIds];
        var foldersToProcess = [...rootFolderIds];
        var parentMap = {}; // childId -> parentId
        var nameMap = {};   // folderId -> folderName

        while (foldersToProcess.length > 0) {
            var childFolders = [];
            var batchSize = 1000;
            for (var i = 0; i < foldersToProcess.length; i += batchSize) {
                var batch = foldersToProcess.slice(i, i + batchSize);
                var folderSearch = search.create({
                    type: 'folder',
                    filters: [
                        ['parent', 'anyof', batch]
                    ],
                    columns: ['internalid', 'name', 'parent']
                });

                folderSearch.run().each(function (result) {
                    var id = String(result.id);
                    var name = result.getValue('name');
                    var parentVal = result.getValue('parent');
                    var parent = parentVal ? String(parentVal) : '';

                    parentMap[id] = parent;
                    nameMap[id] = name;

                    childFolders.push(id);
                    return true;
                });
            }

            if (childFolders.length === 0) {
                break;
            }

            allFolderIds.push(...childFolders);
            foldersToProcess = childFolders;
        }

        // Fetch details for the root folders themselves
        var rootFolderSearch = search.create({
            type: 'folder',
            filters: [
                ['internalid', 'anyof', rootFolderIds]
            ],
            columns: ['internalid', 'name', 'parent']
        });

        rootFolderSearch.run().each(function (result) {
            var id = String(result.id);
            var name = result.getValue('name');
            var parentVal = result.getValue('parent');
            var parent = parentVal ? String(parentVal) : '';

            parentMap[id] = parent;
            nameMap[id] = name;
            return true;
        });

        return {
            allFolderIds: allFolderIds,
            parentMap: parentMap,
            nameMap: nameMap
        };
    }

    /**
     * Climbs up parentMap to determine path segments relative to root folders
     */
    function getRelativeFolderPath(folderId, rootFolderIds, parentMap, nameMap) {
        var idStr = String(folderId);
        if (rootFolderIds.indexOf(idStr) !== -1) {
            return '';
        }

        var pathSegments = [];
        var currentId = idStr;

        while (currentId && rootFolderIds.indexOf(currentId) === -1) {
            var name = nameMap[currentId];
            if (name) {
                pathSegments.unshift(name);
            }
            currentId = parentMap[currentId];
        }

        return pathSegments.join('/');
    }

    return {
        getInputData: getInputData,
        map: map,
        summarize: summarize
    };
});
