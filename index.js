/*
 * spfxrestapi
 * IMPORTANT
 * Santander web _api client framework
 * Dependencies:
 *  - jQuery
 *  
 */

window.spfxrestapi = window.spfxrestapi || new Object();

window.spfxrestapi = {
    Functions: {
        /*
        /// <summary>
        /// Converts the value of objects to strings based on the formats specified and inserts them into another string.
        /// </summary>
        /// <param name="string">the original string</param>
        /// <param name="args[]">relacement args</param>
        /// <returns>
        /// Formated string
        /// </returns>
        */
        stringFormat: function () {
            var s = arguments[0];

            for (var i = 0; i < arguments.length - 1; i++) {
                var reg = new RegExp("\\{" + i + "\\}", "gm");
                s = s.replace(reg, arguments[i + 1]);
            }
            return s;
        },
        /*
       /// <summary>
       /// Gets a value from URL query string by a property key.
       /// </summary>
       /// <param name="name">property key</param>
       /// <returns>
       /// String containing the value for the given property key 
       /// </returns>
       */
        queryString: {
            getByName: function (name) {
                name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
                var regex = new RegExp("[\\?&]" + name + "=([^&#]*)"),
                    results = regex.exec(location.search);
                return results === null ? "" : decodeURIComponent(results[1].replace(/\+/g, " "));
            }
        },
        /*
        /// <summary>
        /// Get an JSON object subset based on a Master JSON and a property value.
        /// </summary>
        /// <param name="JSON">JSON Object</param>
        /// <param name="value">value to look for</param>
        /// <param name="propIn">property key</param>
        /// <returns>
        /// JSON subset object
        /// </returns>
        */
        getJSONObj: function (JSON, value, propIn) {
            var returnValue;
            jQuery.each(JSON, function (_index, _value) {
                if (_value[propIn].toString().toLowerCase() === value.toString().toLowerCase()) {
                    returnValue = _value;
                    return;
                }
            });

            return returnValue;
        },
        bindTemplateElement: function (template, data) {
            return jQuery("#" + template).html().renderHtml(data);
        },
        getMultiScripts: function (arr, path) {
            var _arr = jQuery.map(arr, function (scr) {
                return jQuery.getScript((path || "") + scr);
            });
            _arr.push(jQuery.Deferred(function (deferred) {
                jQuery(deferred.resolve);
            }));
            return jQuery.when.apply(jQuery, _arr);
        },
        executeQueryPromise: function (ctx, result) {
            result = result || {};
            var d = jQuery.Deferred();
            ctx.executeQueryAsync(function () {
                d.resolve(result);
            }, function (sender, args) {
                d.reject(args);
            });
            return d.promise();
        },
        RemoveLastDirectoryPartOf: function (the_url) {
            var the_arr = the_url.split('/');
            the_arr.pop();
            return (the_arr.join('/'));
        }
    },
    SharePoint: {
        Web: {
            /*
            /// <summary>
            /// Get an SPWeb
            /// </summary>
            /// <param name="webAbsoluteUrl">SPWeb absolute URL</param>
            /// <param name="success">success callback function</param>
            /// <param name="failure">error callback function</param>
             /// <param name="properties">array of properties to select from SPWeb</param>
            /// <returns>
            /// SPWeb object
            /// </returns>
            */
            getWebProperties: function (webAbsoluteUrl, success, failure, properties) {

                var select_fields = [];
                //<validation>
                if (success.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::success must be a valid function");
                }
                if (failure.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::failure must be a valid function");
                }
                if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                    throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                }
                if (properties.length === 0) {
                    select_fields = ["Title"];
                } else {
                    select_fields = properties;
                }
                //</validation>

                var rest_filter = "?$select=" + select_fields.join(",");

                var requestUri = webAbsoluteUrl + "/_api/web" + rest_filter;

                jQuery.ajax({
                    url: requestUri,
                    contentType: "application/json;odata=verbose",
                    headers: {
                        "accept": "application/json;odata=verbose"
                    },
                    success: function (data) {
                        success(data);
                    },
                    error: function (error) {
                        failure(error);
                    }
                });
            },
            Lists: {
                /*
                /// <summary>
                /// Delete SPListItem from SPList
                /// </summary>
                /// <param name="webAbsoluteUrl">SPWeb absolute URL</param>
                /// <param name="listId">SPList identifier GUID or List Title will work</param>
                /// <param name="ItemId">SPListItem Id to delete from SPList</param>
                /// <param name="success">success callback function</param>
                /// <param name="failure">error callback function</param>
                /// <returns>
                /// context response through callback function (success/failure)
                /// </returns>
                */
                deleteListItem: function (webAbsoluteUrl, listId, ItemId, success, failure) {

                    //<validation>
                    if (listId.constructor !== String || listId.length === 0) {
                        throw new Error("[[spfxrestapi]]::listId must be a string, and cannot be empty");
                    }
                    if (success.constructor !== Function) {
                        throw new Error("[[spfxrestapi]]::success must be a valid function");
                    }
                    if (failure.constructor !== Function) {
                        throw new Error("[[spfxrestapi]]::failure must be a valid function");
                    }
                    if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                        throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                    }
                    if (ItemId.constructor !== Number || ItemId.length === 0) {
                        throw new Error("[[spfxrestapi]]::itemId must be a Number, and cannot be empty");
                    }
                    //</validation>

                    var digest = "";

                    jQuery.ajax({
                        url: webAbsoluteUrl + "/_api/contextinfo",
                        method: "POST",
                        headers: {
                            "ACCEPT": "application/json;odata=verbose",
                            "content-type": "application/json;odata=verbose"
                        },
                        success: function (data) {
                            digest = data.d.GetContextWebInformation.FormDigestValue;
                        },
                        error: function (error) {
                            console.log(error);
                        }
                    }).done(function () {

                        jQuery.ajax({
                            url: webAbsoluteUrl + "/_api/web/lists/GetByTitle('" + listId + "')/items(" + ItemId + ")",
                            method: "POST",
                            headers: {
                                "X-RequestDigest": digest,
                                "IF-MATCH": "*",
                                "X-HTTP-Method": "DELETE"
                            },
                            success: function (data) {
                                if (typeof (success) !== "undefined") {
                                    success(data);
                                }
                            },
                            error: function (data) {
                                if (typeof (failure) !== "undefined") {
                                    failure(data);
                                }
                            }
                        });
                    });
                },
                /*
                /// <summary>
                /// Delete SPListItem from SPList
                /// </summary>
                /// <param name="webAbsoluteUrl">SPWeb absolute URL</param>
                /// <param name="listId">SPList identifier GUID or List Title will work</param>
                /// <param name="ItemId">SPListItem Id to delete from SPList</param>
                /// <param name="success">success callback function</param>
                /// <param name="failure">error callback function</param>
                /// <returns>
                /// context response through callback function (success/failure)
                /// </returns>
                */
                createListItem: function (webAbsoluteUrl, listId, itemProperties, success, failure) {

                    //<validation>
                    if (listId.constructor !== String || listId.length === 0) {
                        throw new Error("[[spfxrestapi]]::listId must be a string, and cannot be empty");
                    }
                    if (success.constructor !== Function) {
                        throw new Error("[[spfxrestapi]]::success must be a valid function");
                    }
                    if (failure.constructor !== Function) {
                        throw new Error("[[spfxrestapi]]::failure must be a valid function");
                    }
                    if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                        throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                    }
                    //</validation>

                    jQuery.ajax({
                        url: webAbsoluteUrl + "/_api/web/lists/getbytitle('" + listId + "')?$select=ListItemEntityTypeFullName",
                        type: "GET",
                        contentType: "application/json;odata=verbose",
                        headers: {
                            "Accept": "application/json;odata=verbose",
                            "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
                        },
                        success: function (data) {

                            itemProperties["__metadata"] = {
                                "type": data.d.ListItemEntityTypeFullName
                            };

                            jQuery.ajax({
                                url: webAbsoluteUrl + "/_api/web/lists/getbytitle('" + listId + "')/items",
                                type: "POST",
                                contentType: "application/json;odata=verbose",
                                data: JSON.stringify(itemProperties),
                                headers: {
                                    "Accept": "application/json;odata=verbose",
                                    "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
                                },
                                success: function (data) {
                                    if (typeof (success) !== "undefined") {
                                        success(data);
                                    }
                                },
                                error: function (data) {
                                    if (typeof (failure) !== "undefined") {
                                        failure(data);
                                    }
                                }
                            });
                        },
                        error: function (data) {
                            failure(data);
                        }
                    });
                },
                updateListItem: function (webAbsoluteUrl, listTitle, oldItem, newItem, success, failure) {

                    //<validation>
                    if (listTitle.constructor !== String || listTitle.length === 0) {
                        throw new Error("[[spfxrestapi]]::listTitle must be a string, and cannot be empty");
                    }
                    if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                        throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                    }
                    //</validation>

                    newItem["__metadata"] = {
                        "type": oldItem.__metadata.type
                    };

                    jQuery.ajax({
                        url: webAbsoluteUrl + "/_api/web/lists/getbytitle('" + listTitle + "')/items(" + oldItem.ID + ")",
                        type: "POST",
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
                            "content-Type": "application/json;odata=verbose",
                            "X-Http-Method": "PATCH",
                            "If-Match": "*"
                        },
                        data: JSON.stringify(newItem),
                        success: function (data) {
                            if (typeof (success) !== "undefined") {
                                success(data);
                            }
                        },
                        error: function (data) {
                            if (typeof (failure) !== "undefined") {
                                failure(data);
                            }
                        }
                    });
                },
                getListItemsByField: function (webAbsoluteUrl, listId, itemId, field, properties) {

                    var _properties = "";

                    //<validation>
                    if (listId.constructor !== String || listId.length === 0) {
                        throw new Error("[[spfxrestapi]]::listTitle must be a string, and cannot be empty");
                    }

                    if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                        throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                    }

                    if (properties !== undefined) {
                        _properties = "$select=" + properties.join(",") + "&";

                    }

                    //</validation>

                    var regexGuid = /^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$/gi;

                    var isGuid = regexGuid.test(listId);

                    var _call = isGuid ? webAbsoluteUrl + "/_api/web/lists(guid'" + listId + "')/Items" : webAbsoluteUrl + "/_api/web/lists/getbytitle('" + listId + "')/Items?";

                    var filter = "$filter=" + "(" + field + " eq '" + itemId + "')";

                    var _url = _call + _properties + filter;

                    return jQuery.ajax({
                        url: _url,
                        type: "GET",
                        headers: {
                            "accept": "application/json;odata=verbose"
                        }
                    });
                },
                getListItemByQueryText: function (listTitle, querytext) {

                    return jQuery.ajax({
                        url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('" + listTitle + "')/items?$filter=" + querytext,
                        type: "GET",
                        headers: {
                            "accept": "application/json;odata=verbose"
                        }
                    });
                },
                getListItemByFilter: function (webAbsoluteUrl, listTitle, querytext) {

                    return jQuery.ajax({
                        url: webAbsoluteUrl + "/_api/web/lists/getbytitle('" + listTitle + "')/items?" + querytext,
                        type: "GET",
                        headers: {
                            "accept": "application/json;odata=verbose"
                        }
                    });
                },
                getListItemByCaml: function (webAbsoluteUrl, listTitle, CamlQuery) {
                    var queryViewXml = "<View>" + CamlQuery + "</View>";
                    var params = {};
                    params.create = {};
                    params.create.sUrl = webAbsoluteUrl;
                    params.create.lName = listTitle;
                    params.create.filter = "";
                    params.create.body = {
                        'query': {
                            '__metadata': {
                                'type': 'SP.CamlQuery'
                            },
                            'ViewXml': queryViewXml
                        }
                    }

                    var rquest = {
                        method: 'POST',
                        url: params.create.sUrl + "/_api/web/lists/getByTitle('" + params.create.lName + "')/getitems",
                        header: {
                            "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                            "accept": "application/json;odata=verbose",
                            "content-type": "application/json;odata=verbose"
                        },
                        body: params.create.body,
                        contentType: "application/json;odata=verbose",

                    };

                    return jQuery.ajax({
                        method: rquest.method,
                        url: rquest.url,
                        contentType: rquest.contentType,
                        headers: rquest.header,
                        data: JSON.stringify(rquest.body),
                    });
                },
                getListItems: function (webAbsoluteUrl, listId, OnSuccess, OnComplete, filter) {

                    //<validation>
                    if (listId.constructor !== String || listId.length === 0) {
                        throw new Error("[[spfxrestapi]]::listId must be a string, and cannot be empty");
                    }

                    if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                        throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                    }
                    //</validation>

                    var regexGuid = /^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$/gi;

                    var isGuid = regexGuid.test(listId);

                    var url = isGuid ? webAbsoluteUrl + "/_api/web/lists(guid'" + listId + "')/items" + filter : webAbsoluteUrl + "/_api/web/lists/getbytitle('" + listId + "')/items" + filter;

                    function GetItems() {
                        jQuery.ajax({
                            url: url,
                            method: "GET",
                            headers: {
                                "Accept": "application/json;odata=verbose"
                            },
                            success: function (data) {
                                results = results.concat(data.d.results);
                                if (data.d.__next) {
                                    url = data.d.__next;
                                    GetItems();
                                } else {
                                    nonext = true;
                                    if (typeof OnSuccess !== "undefined") {
                                        OnSuccess(results);
                                    }
                                }
                            },
                            error: function (error) {
                                console.log(error);
                            },
                            complete: function () {
                                if (typeof OnComplete !== "undefined" && nonext) {
                                    OnComplete(results);
                                }
                            }
                        });
                    }

                    var nonext = false;
                    var results = [];
                    GetItems();
                },
                getDocumentsFolder: function (webAbsoluteUrl, folder) {

                    //<validation>
                    if (folder.constructor !== String || folder.length === 0) {
                        throw new Error("[[spfxrestapi]]::folder must be a string, and cannot be empty");
                    }

                    if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                        throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                    }
                    //</validation>

                    var _url = webAbsoluteUrl + "/_api/Web/GetFolderByServerRelativeUrl('" + folder + "')/Files?";

                    var selectQuery = "$select=ListItemAllFields/CardImage,ListItemAllFields/Synopsis,ListItemAllFields/FileRef,ListItemAllFields/UniqueId,ListItemAllFields/Title,ListItemAllFields/Name,ListItemAllFields/ServerRelativeUrl&";

                    var expandQuery = "$expand=ListItemAllFields";

                    _url = _url + selectQuery + expandQuery;

                    return jQuery.ajax({
                        url: _url,
                        type: "GET",
                        headers: {
                            "accept": "application/json;odata=verbose"
                        }
                    });
                },
                getListItemFields: function (webAbsoluteUrl, listId) {

                    //<validation>
                    if (listId.constructor !== String || listId.length === 0) {
                        throw new Error("[[spfxrestapi]]::listTitle must be a string, and cannot be empty");
                    }

                    if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                        throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                    }
                    //</validation>

                    var regexGuid = /^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$/gi;

                    var isGuid = regexGuid.test(listId);

                    var _url = isGuid ? webAbsoluteUrl + "/_api/web/lists(guid'" + listId + "')/Fields" : webAbsoluteUrl + "/_api/web/lists/getbytitle('" + listId + "')/Fields";

                    return jQuery.ajax({
                        url: _url,
                        type: "GET",
                        headers: {
                            "accept": "application/json;odata=verbose"
                        }
                    });
                },
                uploadListItemAttachment: function (webAbsoluteUrl, listId, file, ID) {

                    //<validation>
                    if (listId.constructor !== String || listId.length === 0) {
                        throw new Error("[[spfxrestapi]]::listTitle must be a string, and cannot be empty");
                    }

                    if (ID.constructor !== Number || ID.length === 0) {
                        throw new Error("[[spfxrestapi]]::ID must be a number, and cannot be empty");
                    }

                    if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                        throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                    }
                    //</validation>

                    var regexGuid = /^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$/gi;

                    var isGuid = regexGuid.test(listId);

                    var _url = isGuid ? webAbsoluteUrl + "/_api/web/lists(guid'" + listId + "')/Items" : webAbsoluteUrl + "/_api/web/lists/getbytitle('" + listId + "')/Items";

                    var _file = file[0].files[0];

                    getFileBuffer(file)
                        .then(function (buffer) {

                            jQuery.ajax({
                                url: _url + "(" + ID + ")/AttachmentFiles/add(FileName='" + _file.name + "')",
                                method: "POST",
                                data: buffer,
                                processData: false,
                                headers: {
                                    "Accept": "application/json; odata=verbose",
                                    "content-type": "application/json; odata=verbose",
                                    "X-RequestDigest": document.getElementById("__REQUESTDIGEST").value,
                                    "content-length": buffer.byteLength
                                }
                            });
                        });

                    // Get the local file as an array buffer.
                    function getFileBuffer() {
                        var deferred = jQuery.Deferred();
                        var reader = new FileReader();
                        reader.onloadend = function (e) {
                            deferred.resolve(e.target.result);
                        };
                        reader.onerror = function (e) {
                            deferred.reject(e.target.error);
                        };
                        reader.readAsArrayBuffer(_file);
                        return deferred.promise();
                    }
                },
                deleteAttachment: function (webAbsoluteUrl, listTitle, itemId, fileName) {

                    //<validation>
                    if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                        throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                    }
                    if (listTitle.constructor !== String || listTitle.length === 0) {
                        throw new Error("[[spfxrestapi]]::listTitle must be a string, and cannot be empty");
                    }
                    if (itemId.constructor !== String || itemId.length === 0) {
                        throw new Error("[[spfxrestapi]]::itemId must be a string, and cannot be empty");
                    }
                    if (fileName.constructor !== String || fileName.length === 0) {
                        throw new Error("[[spfxrestapi]]::fileName must be a string, and cannot be empty");
                    }
                    //</validation>

                    return jQuery.ajax({
                        url: webAbsoluteUrl + "/_api/lists/getByTitle('" + listTitle + "')/getItemById(" + itemId + ")/AttachmentFiles/getByFileName('" + fileName + "')",
                        method: "DELETE",
                        headers: {
                            "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
                        }
                    });
                }
            }
        },
        Users: {
            getUserGroups: function (webAbsoluteUrl, userId, success, failure) {

                //<validation>
                if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                    throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                }
                if (userId.constructor !== Number || userId.length === 0) {
                    throw new Error("[[spfxrestapi]]::userId must be a number, and cannot be empty");
                }
                if (success.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::success must be a valid function");
                }

                if (failure.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::failure must be a valid function");
                }
                //</validation>

                var requestUri = webAbsoluteUrl + "/_api/web/GetUserById(" + userId + ")/Groups";

                jQuery.ajax({
                    url: requestUri,
                    contentType: "application/json;odata=verbose",
                    headers: {
                        "accept": "application/json;odata=verbose"
                    },
                    success: function (data) {
                        success(data);
                    },
                    error: function (error) {
                        failure(error);
                    }
                });
            },
            userInGroup: function (webAbsoluteUrl, userId, groupTitle, success, failure) {

                //<validation>
                if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                    throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                }
                if (userId.constructor !== Number || userId.length === 0) {
                    throw new Error("[[spfxrestapi]]::userId must be a number, and cannot be empty");
                }
                if (groupTitle.constructor !== String || groupTitle.length === 0) {
                    throw new Error("[[spfxrestapi]]::groupTitle must be a string, and cannot be empty");
                }
                if (success.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::success must be a valid function");
                }

                if (failure.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::failure must be a valid function");
                }
                //</validation>

                spfxrestapi.SharePoint.Users.getUserGroups(webAbsoluteUrl, userId, function (response) {
                    success(response.d.results);
                }, function (error) {
                    failure(error);
                });
            },
            getUserById: function (webAbsoluteUrl, userId, success, failure) {

                //<validation>
                if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                    throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                }
                if (userId.constructor !== Number || userId.length === 0) {
                    throw new Error("[[spfxrestapi]]::userId must be a number, and cannot be empty");
                }
                if (success.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::success must be a valid function");
                }

                if (failure.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::failure must be a valid function");
                }
                //</validation>

                var requestUri = webAbsoluteUrl + "/_api/web/getuserbyid(" + userId + ")";

                jQuery.ajax({
                    url: requestUri,
                    contentType: "application/json;odata=verbose",
                    headers: {
                        "accept": "application/json;odata=verbose"
                    },
                    success: function (data) {
                        success(data);
                    },
                    error: function (error) {
                        failure(error);
                    }
                });
            },
            ensureUser: function (webAbsoluteUrl, loginName, success, failure) {

                //<validation>
                if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                    throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                }
                if (loginName.constructor !== String || loginName.length === 0) {
                    throw new Error("[[spfxrestapi]]::userId must be a string, and cannot be empty");
                }
                if (groupTitle.constructor !== String || groupTitle.length === 0) {
                    throw new Error("[[spfxrestapi]]::loginName must be a string, and cannot be empty");
                }
                if (success.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::success must be a valid function");
                }

                if (failure.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::failure must be a valid function");
                }
                //</validation>

                var payload = {
                    "logonName": loginName
                };

                jQuery.ajax({
                    url: webAbsoluteUrl + "/_api/web/ensureuser",
                    type: "POST",
                    contentType: "application/json;odata=verbose",
                    data: JSON.stringify(payload),
                    headers: {
                        "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
                        "accept": "application/json;odata=verbose"
                    },
                    success: function (data) {
                        success(data);
                    },
                    error: function (error) {
                        failure(error);
                    }
                });
            },
            getUserByEmail: function (webAbsoluteUrl, email, success, failure) {

                //<validation>
                if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                    throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                }
                if (email.constructor !== String || email.length === 0) {
                    throw new Error("[[spfxrestapi]]::email must be a string, and cannot be empty");
                }
                if (success.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::success must be a valid function");
                }

                if (failure.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::failure must be a valid function");
                }
                //</validation>

                jQuery.ajax({
                    url: webAbsoluteUrl + "/_api/Web/SiteUsers?$filter=Email eq '" + email + "'",
                    type: "GET",
                    contentType: "application/json;odata=verbose",
                    headers: {
                        "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
                        "accept": "application/json;odata=verbose"
                    },
                    success: function (data) {
                        success(data);
                    },
                    error: function (error) {
                        failure(error);
                    }
                });
            }
        },
        SOD: {
            executeOrDelay: function (sodScripts, onLoadAction) {
                if (SP.SOD.loadMultiple) {
                    for (var x = 0; x < sodScripts.length; x++) {
                        //register any unregistered scripts
                        if (!_v_dictSod[sodScripts[x]]) {
                            SP.SOD.registerSod(sodScripts[x], "/_layouts/15/" + sodScripts[x]);
                        }
                    }
                    SP.SOD.loadMultiple(sodScripts, onLoadAction);
                } else
                    ExecuteOrDelayUntilScriptLoaded(onLoadAction, sodScripts[0]);
            }
        },
        Navigation: {
            getNavigation: function (webAbsoluteUrl, success, failure, provider) {

                //<validation>
                if (provider.constructor !== String || provider.length === 0) {
                    provider = "GlobalNavigationSwitchableProvider";
                }

                if (success.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::success must be a valid function");
                }

                if (failure.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::failure must be a valid function");
                }
                if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                    throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                }
                //</validation>

                var request = webAbsoluteUrl + "/_api/navigation/menustate?mapprovidername='" + provider + "'";

                return jQuery.ajax({
                    method: "GET",
                    url: request,
                    headers: {
                        "accept": "application/json;odata=verbose"
                    },
                    success: function (data) {

                        var nodes = data.d.MenuState.Nodes.results;

                        var that = this;

                        if (typeof (success) !== "undefined") {
                            success(nodes);
                        }
                    },
                    error: function (error) {
                        if (typeof (failure) !== "undefined") {
                            failure(error);
                        }
                    }
                });
            }
        },
        Taxonomy: {
            createTerm: function (termStoreId, termSetId, termName, callback) {

                var sp_files = [
                    'SP.Runtime.js',
                    'SP.js',
                    'SP.Taxonomy.js'
                ];

                var layoutsPath = _spPageContextInfo.siteAbsoluteUrl + "/_layouts/15/";
                spfxrestapi.Functions.getMultiScripts(sp_files, layoutsPath).done(function () {
                    var ctx = SP.ClientContext.get_current();
                    var taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(ctx);
                    var termStores = taxonomySession.get_termStores();
                    var termSet = termStores.getById(termStoreId).getTermSet(termSetId);
                    var termGuid = SP.Guid.newGuid();
                    var newTerm = termSet.createTerm(termName, 1033, termGuid.toString());
                    newTerm.setDescription(termName, 1033);
                    ctx.load(newTerm);
                    ctx.executeQueryAsync(Function.createDelegate(this, function (sender, args) {
                            callback(sender, args);
                        }),
                        Function.createDelegate(this, function (sender, args) {
                            callback(sender, args);
                        }));
                });
            }
        },
        Utilities: {
            sendMail: function (webAbsoluteUrl, from, to, body, subject, success, failure) {

                //<validation>
                if (success.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::success must be a valid function");
                }
                if (failure.constructor !== Function) {
                    throw new Error("[[spfxrestapi]]::failure must be a valid function");
                }
                if (webAbsoluteUrl.constructor !== String || webAbsoluteUrl.length === 0) {
                    throw new Error("[[spfxrestapi]]::webAbsoluteUrl must be a string, and cannot be empty");
                }
                //</validation>

                var url = webAbsoluteUrl + "/_api/SP.Utilities.Utility.SendEmail";
                jQuery.ajax({
                    contentType: 'application/json',
                    url: url,
                    type: "POST",
                    data: JSON.stringify({
                        'properties': {
                            '__metadata': {
                                'type': 'SP.Utilities.EmailProperties'
                            },
                            'From': from,
                            'To': {
                                'results': [to]
                            },
                            'Body': body,
                            'Subject': subject
                        }
                    }),
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        "content-type": "application/json;odata=verbose",
                        "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
                    },
                    success: function (data) {
                        success();
                    },
                    error: function (err) {
                        failure();
                    }
                });
            }
        }
    }
};