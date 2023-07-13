"use strict";
var Sdk = window.Sdk || {};

/**
 * @function getClientUrl
 * @description Get the client URL.
 * @returns {string} The client URL.
 */
Sdk.getClientUrl = function () {
    var context;
    // GetGlobalContext defined by including reference to
    // ClientGlobalContext.js.aspx in the HTML page.
    if (typeof GetGlobalContext != "undefined") {
        context = GetGlobalContext();
    } else {
        if (typeof Xrm != "undefined") {
            // Xrm.Page.context defined within the Xrm.Page object model for form scripts.
            context = Xrm.Page.context;
        } else {
            throw new Error("Context is not available.");
        }
    }
    return context.getClientUrl();
};

/**
 * An object instantiated to manage detecting the
 * Web API version in conjunction with the
 * Sdk.retrieveVersion function
 */
Sdk.versionManager = new (function () {
    //Start with base version
    var _webAPIMajorVersion = 8;
    var _webAPIMinorVersion = 0;
    //Use properties to increment version and provide WebAPIPath string used by Sdk.request;
    Object.defineProperties(this, {
        WebAPIMajorVersion: {
            get: function () {
                return _webAPIMajorVersion;
            },
            set: function (value) {
                if (typeof value != "number") {
                    throw new Error(
                        "Sdk.versionManager.WebAPIMajorVersion property must be a number."
                    );
                }
                _webAPIMajorVersion = parseInt(value, 10);
            },
        },
        WebAPIMinorVersion: {
            get: function () {
                return _webAPIMinorVersion;
            },
            set: function (value) {
                if (isNaN(value)) {
                    throw new Error(
                        "Sdk.versionManager._webAPIMinorVersion property must be a number."
                    );
                }
                _webAPIMinorVersion = parseInt(value, 10);
            },
        },
        WebAPIPath: {
            get: function () {
                return "/api/data/v" + _webAPIMajorVersion + "." + _webAPIMinorVersion;
            },
        },
    });
})();

//Setting variables specific to this sample within a container so they won't be
// overwritten by another scripts code
Sdk.SampleVariables = {
    entitiesToDelete: [], // Entity URIs to be deleted later (if user so chooses)
    deleteData: true, // Controls whether sample data are deleted at the end of sample run
    contact1Uri: null, // e.g.: Peter Cambel
    contactAltUri: null, // e.g.: Peter_Alt Cambel
    account1Uri: null, // e.g.: Contoso, Ltd
    account2Uri: null, // e.g.: Fourth Coffee
    contact2Uri: null, // e.g.: Susie Curtis
    opportunity1Uri: null, // e.g.: Adventure Works
    competitor1Uri: null,
};

/**
 * @function request
 * @description Generic helper function to handle basic XMLHttpRequest calls.
 * @param {string} action - The request action. String is case-sensitive.
 * @param {string} uri - An absolute or relative URI. Relative URI starts with a "/".
 * @param {object} data - An object representing an entity. Required for create and update actions.
 * @param {object} addHeader - An object with header and value properties to add to the request
 * @returns {Promise} - A Promise that returns either the request object or an error object.
 */
Sdk.request = function (action, uri, data, addHeader) {
    if (!RegExp(action, "g").test("POST PATCH PUT GET DELETE")) {
        // Expected action verbs.
        throw new Error(
            "Sdk.request: action parameter must be one of the following: " +
            "POST, PATCH, PUT, GET, or DELETE."
        );
    }
    if (!typeof uri === "string") {
        throw new Error("Sdk.request: uri parameter must be a string.");
    }
    if (RegExp(action, "g").test("POST PATCH PUT") && !data) {
        throw new Error(
            "Sdk.request: data parameter must not be null for operations that create or modify data."
        );
    }
    if (addHeader) {
        if (
            typeof addHeader.header != "string" ||
            typeof addHeader.value != "string"
        ) {
            throw new Error(
                "Sdk.request: addHeader parameter must have header and value properties that are strings."
            );
        }
    }

    // Construct a fully qualified URI if a relative URI is passed in.
    if (uri.charAt(0) === "/") {
        //This sample will try to use the latest version of the web API as detected by the
        // Sdk.retrieveVersion function.
        uri = Sdk.getClientUrl() + Sdk.versionManager.WebAPIPath + uri;
    }

    return new Promise(function (resolve, reject) {
        var request = new XMLHttpRequest();
        request.open(action, encodeURI(uri), true);
        request.setRequestHeader("OData-MaxVersion", "4.0");
        request.setRequestHeader("OData-Version", "4.0");
        request.setRequestHeader("Accept", "application/json");
        request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        if (addHeader) {
            request.setRequestHeader(addHeader.header, addHeader.value);
        }
        request.onreadystatechange = function () {
            if (this.readyState === 4) {
                request.onreadystatechange = null;
                switch (this.status) {
                    case 200: // Operation success with content returned in response body.
                    case 201: // Create success.
                    case 204: // Operation success with no content returned in response body.
                        resolve(this);
                        break;
                    default: // All other statuses are unexpected so are treated like errors.
                        var error;
                        try {
                            error = JSON.parse(request.response).error;
                        } catch (e) {
                            error = new Error("Unexpected Error");
                        }
                        reject(error);
                        break;
                }
            }
        };
        request.send(JSON.stringify(data));
    });
};

/**
 * @function initializeEntities
 * @description Runs the sample.
 * This sample demonstrates basic CRUD+ operations.
 * Results are sent to the debugger's console window.
 */
Sdk.initializeEntities = function () {

    /**
     * Behavior of this sample varies by version
     * So starting by retrieving the version;
     */

    Sdk.retrieveVersion()
        .then(function () {
            return Sdk.retrieveEntityDefinitions();
        })
        .catch(function (err) {
            console.log("ERROR: " + err.message);
        });
};

Sdk.retrieveVersion = function () {
    return new Promise(function (resolve, reject) {
        Sdk.request("GET", "/RetrieveVersion", null)
            .then(function (request) {
                try {
                    var RetrieveVersionResponse = JSON.parse(request.response);
                    var fullVersion = RetrieveVersionResponse.Version;
                    var versionData = fullVersion.split(".");
                    Sdk.versionManager.WebAPIMajorVersion = parseInt(versionData[0], 10);
                    Sdk.versionManager.WebAPIMinorVersion = parseInt(versionData[1], 10);
                    resolve();
                } catch (err) {
                    reject(new Error("Error processing version: " + err.message));
                }
            })
            .catch(function (err) {
                reject(new Error("Error retrieving version: " + err.message));
            });
    });
};

Sdk.retrieveEntityDefinitions = function () {
    return new Promise(function (resolve, reject) {
        Sdk.request("GET", "/EntityDefinitions", null)
            .then(function (request) {
                // Process response from previous request.
                var response = JSON.parse(request.response);
                console.log("response", response);

                response.value.forEach((item) => {
                    entityList.add({
                        DisplayName: item.DisplayName?.UserLocalizedLabel?.Label ?? "",

                        CollectionSchemaName: item.CollectionSchemaName,
                        CreatedOn: item.CreatedOn,
                        Description: item.Description?.UserLocalizedLabel?.Label ?? "",
                        DisplayCollectionName: item.DisplayCollectionName?.UserLocalizedLabel?.Label ?? "",
                        EntitySetName: item.EntitySetName,
                        ExternalCollectionName: item.ExternalCollectionName,
                        ExternalName: item.ExternalName,
                        IsActivity: item.IsActivity,
                        LogicalCollectionName: item.LogicalCollectionName,
                        LogicalName: item.LogicalName,
                        PrimaryNameAttribute: item.PrimaryNameAttribute,
                        SchemaName: item.SchemaName,
                        TableType: item.TableType
                    });
                });

                Array.from(document.getElementsByClassName("LogicalName")).forEach(element => {
                    element.addEventListener("click", (event) => {
                        window.open("/WebResources/ya_Attributes?data=" + event.target.innerText);
                    }, false)
                });
            })
            .catch(function (err) {
                reject(err);
            });
    });
};

Sdk.retrieveAttributeDefinitions = function (entityLogicalName) {
    return new Promise(function (resolve, reject) {
        Sdk.request("GET", "/EntityDefinitions(LogicalName='" + entityLogicalName + "')/Attributes", null)
            .then(function (request) {
                // Process response from previous request.
                var response = JSON.parse(request.response);
                console.log("response", response);

                response.value.forEach((item) => {
                    attributeList.add({
                        DisplayName: item.DisplayName?.UserLocalizedLabel?.Label ?? "",

                        AttributeType: item.AttributeType,
                        AttributeTypeName: item.AttributeTypeName?.Value ?? "",
                        ColumnNumber: item.ColumnNumber,
                        CreatedOn: item.CreatedOn,
                        DatabaseLength: item.DatabaseLength ?? "",
                        Description: item.Description?.UserLocalizedLabel?.Label ?? "",
                        EntityLogicalName: item.EntityLogicalName,
                        ExternalName: item.ExternalName,
                        FormatName: item.FormatName?.Value ?? "",
                        ImeMode: item.ImeMode ?? "",
                        IsPrimaryName: item.IsPrimaryName,
                        LogicalName: item.LogicalName,
                        MaxLength: item.MaxLength ?? "",
                        SchemaName: item.SchemaName
                    })
                })
            })
            .catch(function (err) {
                reject(err);
            });
    });
};

Sdk.initializeAttributes = function () {
    return new Promise(async function (resolve, reject) {
        await Sdk.retrieveEntities();

        const selected = document.getElementById("entitySelector").value;
        if (selected) {
            return Sdk.retrieveAttributeDefinitions(selected);
        }
    });
};

function getDataParam() {
    var vals = new Array();
    if (location.search != "") {
        vals = location.search.substr(1).split("&");
        for (var i in vals) {
            vals[i] = vals[i].replace(/\+/g, " ").split("=");
        }
        //look for the parameter named 'data'
        var found = false;
        for (var i in vals) {
            if (vals[i][0].toLowerCase() == "data") {
                found = true;
                break;
            }
        }
    }
    if (found) {
        return vals[i][1];
    }
}

Sdk.retrieveEntities = function () {
    return new Promise(function (resolve, reject) {
        Sdk.request("GET", "/EntityDefinitions", null)
            .then(function (request) {
                var response = JSON.parse(request.response);

                const dataParam = getDataParam();

                var entitySelector = document.getElementById("entitySelector");
                response.value.forEach((item) => {
                    var option = document.createElement("option");
                    option.value = item.LogicalName;
                    option.innerHTML = item.DisplayName?.UserLocalizedLabel?.Label ?? item.LogicalName;
                    if (dataParam && dataParam == item.LogicalName) {
                        option.selected = true;
                    }
                    entitySelector.add(option);
                });

                resolve();
            })
            .catch(function (err) {
                reject(err);
            });
    });
};

Sdk.retrieveRelationshipDefinitions = function () {
    return new Promise(function (resolve, reject) {
        Sdk.request("GET", "/RelationshipDefinitions", null)
            .then(function (request) {
                // Process response from previous request.
                var response = JSON.parse(request.response);
                console.log("response", response);

                response.value.forEach((item) => {
                    relationshipList.add({
                        SchemaName: item.SchemaName,

                        IsHierarchical: item.IsHierarchical,
                        ReferencedAttribute: item.ReferencedAttribute,
                        ReferencedEntity: item.ReferencedEntity,
                        ReferencingAttribute: item.ReferencingAttribute,
                        ReferencingEntity: item.ReferencingEntity,
                        RelationshipType: item.RelationshipType
                    })
                })
            })
            .catch(function (err) {
                reject(err);
            });
    });
};
