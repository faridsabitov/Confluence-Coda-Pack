import * as coda from "@codahq/packs-sdk";
export const pack = coda.newPack();

pack.addNetworkDomain("atlassian.com");

// Testing Document: https://coda.io/d/Confluence-Pack_d8BCazGYeZq/Introduction_suPzs#_lun_E
// Extension page: https://developer.atlassian.com/console/myapps/971feb19-447d-4c15-915d-1cdd12c085ab/overview
// Auth: https://auth.atlassian.com/authorize?audience=api.atlassian.com&client_id=zEavHHdhSsFJG0XBFajSIWcogwI4RvNN&scope=write%3Aconfluence-content%20write%3Aconfluence-space%20write%3Aconfluence-file%20read%3Aconfluence-space.summary%20read%3Aconfluence-props%20write%3Aconfluence-props%20manage%3Aconfluence-configuration%20read%3Aconfluence-content.all%20read%3Aconfluence-content.summary%20search%3Aconfluence%20read%3Aconfluence-content.permission%20read%3Aconfluence-user%20write%3Aconfluence-groups%20read%3Aconfluence-groups%20readonly%3Acontent.attachment%3Aconfluence&redirect_uri=https%3A%2F%2Fcoda.io%2FpacksAuth%2Foauth2%2F17572&state=${YOUR_USER_BOUND_VALUE}&response_type=code&prompt=consent

pack.setUserAuthentication({
    type: coda.AuthenticationType.OAuth2,
    // The following two URLs are will be found in the API's documentation.
    authorizationUrl: "https://auth.atlassian.com/authorize",
    tokenUrl: "https://auth.atlassian.com/oauth/token",
    scopes: ["write:confluence-content", "read:confluence-content.all", "write:confluence-content", "read:confluence-content.all", "write:confluence-file", "read:confluence-space.summary", "write:confluence-space", "read:confluence-props", "write:confluence-props", "manage:confluence-configuration", "read:confluence-content.all", "search:confluence", "read:confluence-content.summary", "read:confluence-content.permission", "read:confluence-user", "read:confluence-groups", "write:confluence-groups", "readonly:content.attachment:confluence"],

    // additionalParams: {
    //   audience: "api.atlassian.com",
    //   prompt: "consent",
    // },

    // After approving access, the user should select which instance they want to connect to.
    // requiresEndpointUrl: true,
    // endpointDomain: "atlassian.com",
});


pack.addFormula({
    name: "GetCloudID",
    description: "Getting Cloud ID to start using the pack",
    parameters: [
    ],
    resultType: coda.ValueType.String,
    execute: async function ([], context) {
        let response = await context.fetcher.fetch({
            method: "GET",
            url: "https://api.atlassian.com/oauth/token/accessible-resources",
            cacheTtlSecs: 0,
        });

        return response.body[0].id
    },
});

const AccessibleResources = coda.makeObjectSchema({
    properties: {
        resourceId: { type: coda.ValueType.String },
        url: { type: coda.ValueType.String },
        name: { type: coda.ValueType.String },
        scopes: { type: coda.ValueType.String },
        avatarUrl: { type: coda.ValueType.String, codaType: coda.ValueHintType.ImageReference },
    },
    displayProperty: "name",
    idProperty: "resourceId",
    featuredProperties: ["resourceId", "url"]
});

pack.addSyncTable({
    name: "SyncAccessibleResources",
    description: "Sync all available Atlassian instances to set up further connections",
    identityName: "SyncAccessibleResources",
    schema: AccessibleResources,
    formula: {
        name: "SyncAccessibleResources",
        description: "Sync all available Atlassian instances to set up further connections",
        parameters: [],
        execute: async function ([], context) {
            let url = "https://api.atlassian.com/oauth/token/accessible-resources";
            let response = await context.fetcher.fetch({
                method: "GET",
                url: url,
            });

            let items = response.body;


            let rows = [];
            for (let item of items) {
                let row = {
                    resourceId: item.id,
                    url: item.url,
                    name: item.name,
                    scopes: item.scopes.toString(),
                    avatarUrl: item.avatarUrl,
                };
                rows.push(row);
            }

            return {
                result: rows,
            };
        },
    },
});


const Page = coda.makeObjectSchema({
    properties: {
        pageId: { type: coda.ValueType.String },
        status: { type: coda.ValueType.String },
        type: { type: coda.ValueType.String },
        title: { type: coda.ValueType.String },
        body: { type: coda.ValueType.String },
        version: { type: coda.ValueType.Number },
        url: { type: coda.ValueType.String },
        apiUrl: { type: coda.ValueType.String },
    },
    displayProperty: "title", // Which property above to display by default.
    idProperty: "pageId"
});

pack.addFormula({
    name: "GetPageContent",
    description: "Getting the information out of a specific page by URL",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "ResourceID",
            description: "An ID of our Confluence instance. You can get the Resource ID out of Sync Accessible Resources table",
        }),
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "PageURL",
            description: "A link to your Confluence page. Ex.: https://{{DOMAIN}}.atlassian.net/wiki/spaces/{{SPACEID}}/pages/{{PAGEID}}/{{PAGETITLE}}",
        }),
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "ContentType",
            description: "Provide the format of content that you would like to get. The available options are: 'storage' (by default), 'atlas_doc_format', 'view', 'export_view', 'styled_view', 'dynamic', 'editor2', 'anonymous_export_view'",
            optional: true,
        }),
    ],
    resultType: coda.ValueType.Object,
    schema: Page,
    execute: async function ([resourceId, pageURL, contentType], context) {
        let page = pageURL.split("/pages/")[1].split("/")[0];
        let url = "https://api.atlassian.com/ex/confluence/" + resourceId + "/rest/api/content/" + page + "?expand=body." + contentType + ",version";
        let response = await context.fetcher.fetch({
            method: "GET",
            url: url,
            cacheTtlSecs: 0,
        });


        //'atlas_doc_format', 'view', 'export_view', 'styled_view', 'dynamic', 'editor2', 'anonymous_export_view'

        // let bodyValue = ""
        // switch(contentType) { 
        //   case "storage": { 
        //       bodyValue = response.body.body.storage.value;
        //       break; 
        //   } 
        //   case "atlas_doc_format": { 
        //       bodyValue = response.body.body.atlas_doc_format.value;
        //       break; 
        //   } 
        //   default: { 
        //       bodyValue = response.body.body.storage.value;
        //       break; 
        //   } 
        // } 

        return {
            pageId: response.body.id,
            status: response.body.status,
            title: response.body.title,
            type: response.body.type,
            body: response.body.body.storage.value,
            version: response.body.version.number,
            url: response.body._links.base.concat(response.body._links.webui),
            apiUrl: response.body._links.self
        };
    },
});

pack.addFormula({
    name: "UpdateConfluencePage",
    description: "Update content of a specific page by URL",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "ResourceID",
            description: "An ID of our Confluence instance. You can get the Resource ID out of Sync Accessible Resources table",
        }),
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "PageURL",
            description: "A link to your Confluence page. Ex.: https://{{DOMAIN}}.atlassian.net/wiki/spaces/{{SPACEID}}/pages/{{PAGEID}}/{{PAGETITLE}}",
        }),
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "VersionNumber",
            description: "A version number of the future page. You can get the current version number from GetPageContent() and then add +1 to show the future version",
        }),
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "Title",
            description: "Provide a title for the page to update. You can get the current title of a page from GetPageContent() formula.",
        }),
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "Type",
            description: "Provide a type of the page to update. You can get the type of a page as an example from GetPageContent() formula. It might be a blog post or a page",
        }),
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "Body",
            optional: true,
            description: "Provide a body for the page to update. You can get the current content of a page as an example from GetPageContent() formula.",
        })
    ],
    resultType: coda.ValueType.String,
    isAction: true,
    execute: async function ([resourceId, pageURL, version, title, contentType, body], context) {
        let page = pageURL.split("/pages/")[1].split("/")[0];
        let url = "https://api.atlassian.com/ex/confluence/" + resourceId + "/rest/api/content/" + page;

        let storageValue: string;
        storageValue = formatHtml(body.toString())

        let response = await context.fetcher.fetch({
            method: "PUT",
            url: url,
            headers: { "Content-Type": "application/json" },
            body: '{"version": { "number": ' + version + '},"title": "' + title + '","type": "' + contentType + '","body": { "storage": { "value":"' + storageValue + '","representation": "storage"}}}'
        });

        return response.toString();
    },
});

type PageInformation = {
    contentType: string,
    version: number
};

pack.addFormula({
    name: "UpdateConfluenceWithCodaPage",
    description: "Export the whole Coda page into Confluence",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "ResourceID",
            description: "An ID of our Confluence instance. You can get the Resource ID out of Sync Accessible Resources table",
        }),
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "PageURL",
            description: "A link to your Confluence page. Ex.: https://{{DOMAIN}}.atlassian.net/wiki/spaces/{{SPACEID}}/pages/{{PAGEID}}/{{PAGETITLE}}",
        }),
        coda.makeParameter({
            type: coda.ParameterType.Html,
            name: "Page",
            description: "Provide a body for the page to update. You can get the current content of a page as an example from GetPageContent() formula.",
        }),
    ],
    resultType: coda.ValueType.String,
    isAction: true,
    execute: async function ([resourceId, pageURL, html], context) {
        let page = pageURL.split("/pages/")[1].split("/")[0];
        let url = "https://api.atlassian.com/ex/confluence/" + resourceId + "/rest/api/content/" + page;

        const pageInformation = await fetchVersion(resourceId, pageURL, context)

        let storageValue: string;
        storageValue = formatHtml(html.toString())

        let pageTitle: string;
        pageTitle = storageValue.split("</h1>")[0].split("<h1>")[1].toString()

        let pageContent: string;
        let cropAmount = pageTitle.length + 9;
        pageContent = storageValue.slice(cropAmount)

        let response = await context.fetcher.fetch({
            method: "PUT",
            url: url,
            headers: { "Content-Type": "application/json" },
            body: '{"version": { "number": ' + pageInformation.version + '},"title": "' + pageTitle + '","type": "' + pageInformation.contentType + '","body": { "storage": { "value":"' + pageContent + '","representation": "storage"}}}'
        });

        return response.toString();
    },
});


async function fetchVersion(resourceId: string, pageURL: string, context): Promise<PageInformation> {
    let page = pageURL.split("/pages/")[1].split("/")[0];
    let url = "https://api.atlassian.com/ex/confluence/" + resourceId + "/rest/api/content/" + page + "?expand=body.storage,version";
    let response = await context.fetcher.fetch({
        method: "GET",
        url: url,
        cacheTtlSecs: 0,
    });

    return {
        contentType: response.body.type,
        version: response.body.version.number + 1
    };
}

pack.addFormula({
    name: "GetHTML",
    description: "<Help text for the formula>",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.Html,
            name: "HTML",
            description: "<Help text for the parameter>",
        }),
        // Add more parameters here and in the array below.
    ],
    resultType: coda.ValueType.String,
    execute: async function ([param], context) {
        return param.toString();
    },
});

pack.addFormula({
    name: "GetFormattedHTML",
    description: "<Help text for the formula>",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.Html,
            name: "HTML",
            description: "<Help text for the parameter>",
        }),
        // Add more parameters here and in the array below.
    ],
    resultType: coda.ValueType.String,
    execute: async function ([param], context) {
        let htmlString = param.toString()
        return formatHtml(htmlString)
    },
});

pack.addFormula({
    name: "GetFormattedHTMLSection",
    description: "<Help text for the formula>",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.Html,
            name: "HTML",
            description: "<Help text for the parameter>",
        }),
        // Add more parameters here and in the array below.
    ],
    resultType: coda.ValueType.String,
    execute: async function ([param], context) {
        let htmlString = param.toString()
        htmlString = formatHtml(htmlString)

        let pageTitle: string;
        // pageContent = htmlString.split("</h1>")[0].toString()
        pageTitle = htmlString.split("</h1>")[0].split("<h1>")[1].toString()

        let cropAmount = pageTitle.length + 9;
        // let cropAmount = 21;
        let pageContent: string = htmlString.slice(cropAmount)


        return pageContent
    },
});

pack.addFormula({
    name: "TransformUser",
    description: "Transform User ID into Confluence syntax that will allow you to tag specific person in the Page Body",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "UserID",
            description: "Get User ID by clicking on the User Profile in Confluence. Ex.: Profile link (https://faridsabitov.atlassian.net/wiki/people/5a55b8ac15a3cc0521fa4da3) -> User ID is 5a55b8ac15a3cc0521fa4da3",
        }),
    ],
    resultType: coda.ValueType.String,
    execute: async function ([userID], context) {
        return "<ac:link><ri:user ri:account-id='" + userID + "' /></ac:link>";
    },
});

let innerHtmlExample = '<span>Update confluence with coda page</span> <table> <h2>Team for Confluence</h2> <thead> <tr> <th>Team Name</th> <th>Management</th> <th>[Team Lead] Dev Manager</th> <th>Tech Lead</th> <th>Developers</th> <th>QA</th> </tr> </thead> <tbody> <tr> <td>Terminators</td> <td><span><a href="https://coda.io/d/_d3e25eOmtM_#People_tuKxc/r3102&amp;view=modal">Liat Refael</a>,<a href="https://coda.io/d/_d3e25eOmtM_#People_tuKxc/r1721&amp;view=modal">Ami Ries</a></span></td> <td><span><a href="https://coda.io/d/_d3e25eOmtM_#People_tuKxc/r3102&amp;view=modal">Liat Refael</a></span> </td> <td><span><a href="https://coda.io/d/_d3e25eOmtM_#People_tuKxc/r2028&amp;view=modal">Eli Abraham</a></span> </td> <td><span><a href="https://coda.io/d/_d3e25eOmtM_#People_tuKxc/r780&amp;view=modal">Tetiana Hrebeniuk</a>,<a href="https://coda.io/d/_d3e25eOmtM_#People_tuKxc/r855&amp;view=modal">Alin Varzaru</a>,<a href="https://coda.io/d/_d3e25eOmtM_#People_tuKxc/r1195&amp;view=modal">Kostiantyn Ivanov</a>,<a href="https://coda.io/d/_d3e25eOmtM_#People_tuKxc/r2028&amp;view=modal">Eli Abraham</a>,<a href="https://coda.io/d/_d3e25eOmtM_#People_tuKxc/r2363&amp;view=modal">Valerii Monakhov</a>,<a href="https://coda.io/d/_d3e25eOmtM_#People_tuKxc/r3257&amp;view=modal">Nick Alistratenko</a></span></td> <td><span><a href="https://coda.io/d/_d3e25eOmtM_#People_tuKxc/r1331&amp;view=modal">Amit Kremnitzer</a>,<a href="https://coda.io/d/_d3e25eOmtM_#People_tuKxc/r1873&amp;view=modal">Ilya Zhdanov</a>,<a href="https://coda.io/d/_d3e25eOmtM_#People_tuKxc/r2242&amp;view=modal">Yan Trunov</a></span> </td> </tr> </tbody> </table>'

pack.addFormula({
    name: "Test",
    description: "",
    parameters: [
    ],
    resultType: coda.ValueType.String,
    execute: function () {
        return formatHtml(innerHtmlExample)
    }
});

function formatHtml(htmlString) {
    let htmlValue = htmlString
    // if htmlValue has table tags inside
    if (htmlValue.includes("<table>")) {
        // Get all tables from htmlValue
        let tables = htmlValue.match(/<table>(.*?)<\/table>/g)
        // Replace each table content with formatted table
        tables.forEach(table => {
            htmlValue = htmlValue.replace(table, formatTable(table))
        })
    }
    if (htmlValue.length !== 0) {
        htmlValue = htmlValue.replace(/"/g, "'");
        htmlValue = htmlValue.replace(/(<span style="font-style: bold;">)(.*?)(<\/span>)/g, "<strong>$2</strong>"); // Replace with strong
        htmlValue = htmlValue.replace(/(<span style="display: inline-block;">)(.*?)(<\/span>)/g, ""); // Remove buttons
        htmlValue = htmlValue.replace(/(<span style="font-style: italic;">)(.*?)(<\/span>)/g, "<em>$2</em>"); // Replace with italic
        // htmlValue = htmlValue.replace(/(<span style="text-decoration: line-through;">)(.*?)(<\/span>)/g,"<em>$2</em>"); // Replace with strikethrough
        htmlValue = htmlValue.replace(/(<span style="text-decoration: underline;">)(.*?)(<\/span>)/g, "<u>$2</u>"); // Replace with underline
        htmlValue = htmlValue.replace(/(<span style="font-family: monospace;">)(.*?)(<\/span>)/g, "<code>$2</code>"); // Replace with monospace code

        // htmlValue = htmlValue.replace(/ background-color: rgb'/g, "color: rgb"); // Replace background color with color only
        // htmlValue = htmlValue.replace(/(margin[^:]*: [^"]*)/g, ""); // Remove margin styles
        // htmlValue = htmlValue.replace(/ style="[^"]*(font-size: [^;]*;)[^"]*"/g,""); // Remove font size styles
        htmlValue = htmlValue.replace(/ style='[^\']*'/g, ""); // Need to identify all styles that we need to remove. Background style is ok to keep //<td data-highlight-colour="#e3fcef">
        htmlValue = htmlValue.replace(/<\/?div[^>]*>/g, "");
        htmlValue = htmlValue.replace(/(<table[^>]*>)(.*?)(<\/table>)/g, "<table>$2</table>");
        htmlValue = htmlValue.replace(/<\/?img[^>]*>/g, "");
        htmlValue = htmlValue.replace(/<\/?br[^>]*>/g, "");
    }
    return htmlValue;
}



// write regex code to delete only the first headings from table tag in html
function formatTable(htmlString) {
    let table = htmlString
    // Find if there are any headings before the tableHead tag inside table tag and move them before the table tag
    let headings = table.match(/<h[1-3][^>]*>.*?<\/h[1-3]>/g)
    if (headings) {
        headings.forEach(heading => {
            table = table.replace(heading, "")
            table = heading + table
        })
    }

    return table
}





// // Replace all inline styles with empty strings
// str = str.replace(/style=".*?"/g, '');

// // Replace all span tags with empty strings
// str = str.replace(/<\/?span>/g, '');

// // Replace all br tags with newline characters
// str = str.replace(/<br>/g, '\n');




// Add support for table title



const Space = coda.makeObjectSchema({
    properties: {
        spaceId: { type: coda.ValueType.String },
        key: { type: coda.ValueType.String },
        name: { type: coda.ValueType.String },
        spaceType: { type: coda.ValueType.String },
        status: { type: coda.ValueType.String },
        homepageLink: { type: coda.ValueType.String },
        webLink: { type: coda.ValueType.String },
        apiLink: { type: coda.ValueType.String },
        // Add more properties here.
    },
    displayProperty: "name", // Which property above to display by default.
    idProperty: "key", // Which property above is a unique ID.
    featuredProperties: ["spaceId", "webLink"]
});



pack.addSyncTable({
    name: "SyncSpaces",
    description: "Get the list of all spaces that you have access to",
    identityName: "SyncSpaces",
    schema: Space,
    formula: {
        name: "SyncSpaces",
        description: "Get the list of all spaces that you have access to",
        parameters: [
            coda.makeParameter({
                type: coda.ParameterType.String,
                name: "ResourceID",
                description: "An ID of our Confluence instance. You can get the Resource ID out of Sync Accessible Resources table",
            }),
            // Add more parameters here and in the array below.
        ],
        execute: async function ([resourceId], context) {
            let start: number = (context.sync.continuation?.index as number) || 0;
            // let isNext: boolean = (context.sync.continuation?.isNext as boolean) || false;
            let url = "https://api.atlassian.com/ex/confluence/" + resourceId + "/rest/api/space?maxResults=100&startAt=";;
            let response = await context.fetcher.fetch({
                method: "GET",
                url: url,
                headers: { "Content-Type": "application/json" },
            });

            let results = response.body.results;

            let rows = [];
            for (let result of results) {
                let row = {
                    spaceId: result.id,
                    key: result.key,
                    name: result.name,
                    spaceType: result.type,
                    status: result.status,
                    homepageLink: result._expandable.homepage,
                    webLink: response.body._links.base + result._links.webui,
                    apiLink: result._links.self,
                };
                rows.push(row);
            }

            start += 25;

            let urlIsNext = response.body._links.next


            let continuation;
            if (start <= 200) {
                continuation = {
                    start: start,
                    isNext: false,
                    continuation: continuation,
                };
            }

            return {
                result: rows,
                continuation: continuation,
            };
        },
    },
});
