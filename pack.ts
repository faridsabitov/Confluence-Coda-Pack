import * as coda from "@codahq/packs-sdk";
export const pack = coda.newPack();



pack.setUserAuthentication({
    type: coda.AuthenticationType.OAuth2,
    authorizationUrl: "https://auth.atlassian.com/authorize",
    tokenUrl: "https://auth.atlassian.com/oauth/token",
    scopes: ["write:confluence-content", "read:confluence-content.all", "read:confluence-content.summary","read:confluence-space.summary", "read:confluence-props", "write:confluence-props", "read:confluence-user"],
    
    //Trying to resolve 60 mins token expiration from Eric's example
    // additionalParams: {
    //     audience: "api.atlassian.com",
    //     prompt: "consent",
    // },
    // // After approving access, the user should select which instance they want to
    // // connect to.
    // requiresEndpointUrl: true,
    // endpointDomain: "atlassian.com",
    // postSetup: [{
    //     type: coda.PostSetupType.SetEndpoint,
    //     name: "SelectEndpoint",
    //     description: "Select the site to connect to:",
    //     // Determine the list of sites they have access to.
    //     getOptions: async function (context) {
    //         let url = "https://api.atlassian.com/oauth/token/accessible-resources";
    //         let response = await context.fetcher.fetch({
    //             method: "GET",
    //             url: url,
    //         });
    //         let sites = response.body;
    //         return sites.map(site => {
    //             // Constructing an endpoint URL from the site ID.
    //             let url = "https://api.atlassian.com/ex/confluence/" + site.id;
    //             return { display: site.name, value: url };
    //         });
    //     },
    // }],

    // // Determines the display name of the connected account.
    // getConnectionName: async function (context) {
    //     // This function is run twice: once before the site has been selected and
    //     // again after. When the site hasn't been selected yet, return a generic
    //     // name.
    //     if (!context.endpoint) {
    //         return "Confluence";
    //     }
    //     // Include both the name of the user and server.
    //     let server = await getServer(context);
    //     let user = await getUser(context);
    //     return `${user.displayName} (${server.serverTitle})`;
    // },
});

// Get information about the Jira server.
async function getServer(context: coda.ExecutionContext) {
    let url = "/rest/api/user/current";
    let response = await context.fetcher.fetch({
        method: "GET",
        url: url,
    });
    return response.body;
}

// Get information about the Jira user.
async function getUser(context: coda.ExecutionContext) {
    let url = "/rest/api/3/myself";
    let response = await context.fetcher.fetch({
        method: "GET",
        url: url,
    });
    return response.body;
}

pack.addNetworkDomain("atlassian.com");

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
    displayProperty: "title",
    idProperty: "pageId"
});

pack.addFormula({
    name: "GetConfluencePage",
    description: "Getting the information out of a specific Confluence page by URL",
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
    execute: async function ([resourceId, pageURL, contentType = "storage"], context) {
        let page = pageURL.split("/pages/")[1].split("/")[0];
        let url = "https://api.atlassian.com/ex/confluence/" + resourceId + "/rest/api/content/" + page + "?expand=body." + contentType + ",version";
        let response = await context.fetcher.fetch({
            method: "GET",
            url: url,
            cacheTtlSecs: 0,
        });

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
    description: "Update content of a specific Confluence page by URL",
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
            name: "Title",
            description: "Provide a title for the page to update. You can get the current title of a page from GetPageContent() formula.",
        }),
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "Body",
            description: "Provide a body for the page to update. You can get the current content of a page as an example from GetPageContent() formula.",
        }),
        coda.makeParameter({
            type: coda.ParameterType.Boolean,
            name: "RemoveCodaLinks",
            description: "Set as true if you want to remove all Coda links from exported HTML",
            suggestedValue: false,
            optional: true
        })
    ],
    resultType: coda.ValueType.String,
    isAction: true,
    execute: async function ([resourceId, pageURL, title, body, removeCodaLinks], context) {
        let page = pageURL.split("/pages/")[1].split("/")[0];
        let url = "https://api.atlassian.com/ex/confluence/" + resourceId + "/rest/api/content/" + page;

        const pageInformation = await fetchVersion(resourceId, pageURL, context)

        let storageValue: string;
        storageValue = formatHtml(body.toString(), removeCodaLinks)

        // Check if title is empty
        if (title === "") { title = "Blank Title" }

        let response = await context.fetcher.fetch({
            method: "PUT",
            url: url,
            headers: { "Content-Type": "application/json" },
            body: '{"version": { "number": ' + pageInformation.version + '},"title": "' + title + '","type": "' + pageInformation.contentType + '","body": { "storage": { "value":"' + storageValue + '","representation": "storage"}}}'
        });

        return statusMessage(response.status.toString());
    },
});

function statusMessage(status: string) {
    switch (status) {
        case "200": return "Page updated successfully";
        case "400": return "Bad request";
        case "401": return "Unauthorized";
        case "403": return "Forbidden";
        case "404": return "Page not found";
        case "500": return "Internal server error";
        case "502": return "Bad gateway";
        case "503": return "Service unavailable";
        default: return "Unknown error";
    }
}

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
        coda.makeParameter({
            type: coda.ParameterType.Boolean,
            name: "RemoveCodaLinks",
            description: "Set as true if you want to remove all Coda links from exported HTML",
            suggestedValue: false,
            optional: true
        })
    ],
    resultType: coda.ValueType.String,
    isAction: true,
    execute: async function ([resourceId, pageURL, html, removeCodaLinks], context) {
        let page = pageURL.split("/pages/")[1].split("/")[0];
        let url = "https://api.atlassian.com/ex/confluence/" + resourceId + "/rest/api/content/" + page;

        const pageInformation = await fetchVersion(resourceId, pageURL, context)

        let storageValue: string;
        storageValue = formatHtml(html.toString(), removeCodaLinks)

        let pageTitle: string;
        pageTitle = storageValue.split("</h1>")[0].split("<h1>")[1].toString()

        let pageContent: string;
        let pageContentLeftPadding: number;
        pageContentLeftPadding = storageValue.split("</h1>")[0].length + "</h1>".length
        pageContent = storageValue.substring(pageContentLeftPadding)   

        let response = await context.fetcher.fetch({
            method: "PUT",
            url: url,
            cacheTtlSecs: 0,
            headers: { "Content-Type": "application/json" },
            body: '{"version": { "number": ' + pageInformation.version + '},"title": "' + pageTitle + '","type": "' + pageInformation.contentType + '","body": { "storage": { "value":"' + pageContent + '","representation": "storage"}}}'
        });

        return statusMessage(response.status.toString());
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
    name: "GetCodaPage",
    description: "Get the code version of Coda page",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.Html,
            name: "CodaPage",
            description: "Provide a name of the Coda page to export",
        }),
    ],
    connectionRequirement: coda.ConnectionRequirement.None,
    resultType: coda.ValueType.String,
    execute: async function ([codaPage], context) {
        return codaPage.toString();
    },
});

pack.addFormula({
    name: "GetCodaPageFormattedForConfluence",
    description: "Get the code version of Coda page formatted for Confluence. You can use it to update the page in Confluence by using 'UpdateConfluencePage' button",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.Html,
            name: "CodaPage",
            description: "Provide a name of the Coda page to export",
        }),
        coda.makeParameter({
            type: coda.ParameterType.Boolean,
            name: "RemoveCodaLinks",
            description: "Set as true if you want to remove all Coda links from exported HTML",
            optional: true
        })
    ],
    resultType: coda.ValueType.String,
    connectionRequirement: coda.ConnectionRequirement.None,
    execute: async function ([codaPage, removeCodaLinks = false], context) {
        let htmlString = codaPage.toString()

        let storageValue: string;
        storageValue = formatHtml(htmlString, removeCodaLinks)

        let pageTitle: string;
        pageTitle = storageValue.split("</h1>")[0].split("<h1>")[1].toString()

        let pageContent: string;
        let pageContentLeftPadding: number;
        pageContentLeftPadding = storageValue.split("</h1>")[0].length + "</h1>".length
        pageContent = storageValue.substring(pageContentLeftPadding)   
        return pageContent;
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

function formatHtml(htmlString, removeCodaLinks = false) {
    let htmlValue = htmlString

    if (htmlValue.length !== 0) {

        //
        // Pre-clean up
        //


        // Remove margin-top: 10px from H1. Coda has first H1 with margin-top: 10px without ";" symbol at the end
        htmlValue = htmlValue.replace(/<h1 style="font-size: 36px; margin-top: 10px">/g, "<h1>");

        // Remove all margin styles
        htmlValue = htmlValue.replace(/margin.*?;/g, ""); 

        // Remove all font size styles
        htmlValue = htmlValue.replace(/font-size.*?;/g, "");

        // Remove all empty space characters from style parameters
        htmlValue = htmlValue.replace(/style=".*?"/g, function (match) {
            return match.replace(/\s/g, "");
        });

        // Remove empty styles
        htmlValue = htmlValue.replace(/ style=""/g, "");

        // Remove div tags without any styles
        // htmlValue = htmlValue.replace(/<\/?div[^>]*>/g, "");


        //
        // Translators
        //
        htmlValue = htmlValue.replace(/"/g, "'"); // Change the double quotes to single quotes
        htmlValue = htmlValue.replace(/(\r\n|\n|\r)/gm, ""); // Remove all new lines

        // Header Translators: https://coda.io/d/Internal-Confluence-Pack_dZqveeLvxRm/Testing_suSQa?playModeWorkflowId=#Testing-Scenarios_tuIvF/r1
        htmlValue = htmlValue.replace(/(<h1>)(.*?)(<\/h1>)/g, "<h1>$2</h1>"); // Header1
        htmlValue = htmlValue.replace(/(<h2>)(.*?)(<\/h2>)/g, "<h2>$2</h2>"); // Header 2
        htmlValue = htmlValue.replace(/(<h3>)(.*?)(<\/h3>)/g, "<h3>$2</h3>"); // Header 3

        // Text effect Translators: https://coda.io/d/Internal-Confluence-Pack_dZqveeLvxRm/Testing_suSQa?playModeWorkflowId=#Testing-Scenarios_tuIvF/r5
        htmlValue = htmlValue.replace(/(<span style='font-weight:bold;'>)(.*?)(<\/span>)/g, "<strong>$2</strong>"); // Replace with strong
        htmlValue = htmlValue.replace(/(<span style='font-style:italic;'>)(.*?)(<\/span>)/g, "<em>$2</em>"); // Replace with italic/emphasis
        htmlValue = htmlValue.replace(/(<span style='text-decoration:line-through;'>)(.*?)(<\/span>)/g, "<del>$2</del>"); // Replace with strikethrough
        htmlValue = htmlValue.replace(/(<span style='text-decoration:underline;'>)(.*?)(<\/span>)/g, "<u>$2</u>"); // Replace with underline
        htmlValue = htmlValue.replace(/(<span style='font-family:monospace;'>)(.*?)(<\/span>)/g, "<code>$2</code>"); // Replace with monospace code
        htmlValue = htmlValue.replace(/(<blockquote[^>]*><span>)(.*?)(<\/span><\/blockquote>)/g, "<blockquote>$2</blockquote>"); // Replace with block quote

        // Text breaks Translators
        // replace all paragraphs
        htmlValue = htmlValue.replace(/<br>/g, "<br />"); // replace all <br> tags with <br /> tags
        htmlValue = htmlValue.replace(/<hr>/g, "<hr />"); // replace all <hr> tags with <hr /> tags

        // List Translators
        htmlValue = htmlValue.replace(/(<ul[^>]*>)(.*?)(<\/ul>)/g, "<ul>$2</ul>"); // remove all parameters from ul tags
        htmlValue = htmlValue.replace(/(<li[^>]*>)(.*?)(<\/li>)/g, "<li>$2</li>"); // remove all parameters from li tags
        htmlValue = htmlValue.replace(/(<ol[^>]*>)(.*?)(<\/ol>)/g, "<ol>$2</ol>"); // remove all parameters from ol tags
        // task list translators
        // start with each item first and then replace them with task code
        //let codaCode = '<ul style="margin-block-start: 1em; margin-block-end: 1em;"><div style="text-align: left; margin-top: 0.5em; margin-bottom: 0.5em;"><input type="checkbox" readonly="" style="width: 2em;"><span>task list item</span></div><div style="text-align: left; margin-top: 0.5em; margin-bottom: 0.5em;"><input type="checkbox" readonly="" style="width: 2em;" checked="true"><span>completed item</span></div></ul>'
        //let confluenceCode = '<ac:task-list><ac:task><ac:task-status>incomplete</ac:task-status><ac:task-body>task list item</ac:task-body></ac:task><ac:task><ac:task-status>complete</ac:task-status><ac:task-body>completed item</ac:task-body></ac:task></ac:task-list>'

        // htmlValue = htmlValue.replace(/<span style="display: inline-block;">(.*?)<\/span>/g, ""); // Remove buttons // Test with icons and larger style buttons as well

        htmlValue = htmlValue.replace(/<caption[^>]*>(.*?)<\/caption>/g, "$1"); // Remove caption tags

        // Text alignment Translators
        // Replace div tags that have text-align styles with p tags
        htmlValue = htmlValue.replace(/(<div style='text-align:left;'>)(.*?)(<\/div>)/g, "<p style='text-align: left;'>$2</p>");
        htmlValue = htmlValue.replace(/(<div style='text-align:right;'>)(.*?)(<\/div>)/g, "<p style='text-align: right;'>$2</p>");

        // htmlValue = htmlValue.replace(/ background-color: rgb'/g, "color: rgb"); // Replace background color with color only
        // htmlValue = htmlValue.replace(/(margin[^:]*: [^"]*)/g, ""); // Remove margin styles
        // htmlValue = htmlValue.replace(/ style="[^"]*(font-size: [^;]*;)[^"]*"/g,""); // Remove font size styles
        // htmlValue = htmlValue.replace(/ style='[^\']*'/g, ""); // Need to identify all styles that we need to remove. Background style is ok to keep //<td data-highlight-colour="#e3fcef">
        // htmlValue = htmlValue.replace(/<\/?div[^>]*>/g, "");
        htmlValue = htmlValue.replace(/<\/?img[^>]*>/g, "");

        // If htmlValue has table tags inside with any paramaters
        if (htmlValue.includes("<table")) {
            // Get all tables from htmlValue
            let tables = htmlValue.match(/<table(.*?)<\/table>/g)

            // Replace each table content with formatted table
            tables.forEach(table => {
                htmlValue = htmlValue.replace(table, formatTable(table))
            })
        }

        // Remove all links related to coda.io if isCodaLinks is true
        if (removeCodaLinks) {
            htmlValue = htmlValue.replace(/<a[^>]*href='https:\/\/coda.io[^>]*>(.*?)<\/a>/g, "$1");
        }

        // Cleaning up the htmlValue
        // Remove all span tags without any parameters
        // htmlValue = htmlValue.replace(/<span>(.*?)<\/span>/g, "$1");

        // Remove all div tags without any parameters
        htmlValue = htmlValue.replace(/<div>(.*?)<\/div>/g, "$1");
    }
    return htmlValue;
}

let innerHtmlExample = '<div><img src="https://cdn.coda.io/icons/png/color/checked-120.png" width="40px" height="40px"/><h1 style="font-size: 36px; margin-top: 10px">Text breaks Translators</h1><div style="text-align: right; margin-top: 0.5em; margin-bottom: 0.5em;"><span>Some text here</span></div><div style="text-align: left; margin-top: 0.5em; margin-bottom: 0.5em;"><span>and then another one in a new line</span><br><span>and last one with CMD+Enter option</span></div><div style="text-align: left; margin-top: 0.5em; margin-bottom: 0.5em;"><br></div></div>'

pack.addFormula({
    name: "Test",
    description: "",
    parameters: [
    ],
    resultType: coda.ValueType.String,
    execute: function () {
        return formatHtml(innerHtmlExample, true)
    }
});


function formatTable(htmlString) {
    // Clean the table tag and assign the value
    let table = htmlString.replace(/(<table[^>]*>)(.*?)(<\/table>)/g, "<table data-layout='wide'>$2</table>");

    // Find if there are any headings before the tableHead tag inside table tag and move them before the table tag
    let headings = table.match(/<h[1-3][^>]*>.*?<\/h[1-3]>/g)
    if (headings) {
        headings.forEach(heading => {
            table = table.replace(heading, "")
            table = heading + table
        })
    }

    // remove all paramaters from all th tags but keep the content
    table = table.replace(/(<th [^>]*>)(.*?)(<\/th>)/g, "<th>$2</th>");
    return table
}

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
    },
    displayProperty: "name",
    idProperty: "key",
    featuredProperties: ["spaceId", "webLink"]
});
