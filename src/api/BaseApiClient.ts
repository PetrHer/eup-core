import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import IApiResult from "./IApiResult";
import { delay } from '../utils/utils';

export abstract class BaseApiClient {

    constructor(protected webPartContext: WebPartContext, protected getUrl: () => string) { }

    /**
     * Retrieves a fresh request digest value from SharePoint for authentication.
     * @returns Promise resolving to the FormDigestValue string.
     */
    protected async getFreshDigest(): Promise<string> {
        const baseUrl = this.getUrl();
        const apiUrl = `${baseUrl}/_api/contextinfo`;
        const response = await this.spPost(apiUrl);
        return response.data && (response.data as any).FormDigestValue ? (response.data as any).FormDigestValue : '';
    }
    /**
     * Sends a POST request to a specified SharePoint API endpoint with the given request body and headers.
     * @param url The URL of the SharePoint API endpoint.
     * @param body The body of the request, which can be an object, string, or ArrayBuffer (default is an empty object).
     * @param headers Additional headers to include in the request (default is an empty object).
     * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data or errors.
     */
    public async spPost(url: string, body: object | string | ArrayBuffer = {}, headers: object = {}): Promise<IApiResult> {
        try {

            const isBatch = (headers as any)['Content-Type']?.startsWith('multipart/mixed');

            const hasCustomContentType =
                Object.keys(headers).some(
                    h => h.toLowerCase() === "content-type"
                );
            const digest = isBatch
                ? await this.getFreshDigest()
                : undefined;
            const spOpts: ISPHttpClientOptions = {
                headers: {
                    ...headers,
                    ...(!isBatch && !hasCustomContentType && {
                        'Accept': 'application/json',
                        'Content-type': 'application/json',
                    }),
                    ...(isBatch && {
                        'X-RequestDigest': digest
                    }),
                },
                body: (typeof body === 'object' && !(body instanceof ArrayBuffer)) ? JSON.stringify(body) : body
            };
            const response: SPHttpClientResponse = await this.webPartContext.spHttpClient.post(url, SPHttpClient.configurations.v1, spOpts);
            const responseJson: any = response.status === 204
                ? true
                : (response.status === 200 && isBatch) ? await response.text() : await response.json();

            const result: IApiResult = {
                ok: response.ok,
                code: response.status
            };

            if (response.ok) {
                result.data = responseJson;
            } else {
                console.error(responseJson);
                result.exception = responseJson;
            }

            return result;
        }
        catch (exception: any) {
            console.error('Error in POST request: ', url, body);
            console.error(exception.error, exception);

            return {
                ok: false,
                code: 500,
                error: exception.message,
                exception
            };
        }
    }

    /**
     * Sends a DELETE request to a specified SharePoint API endpoint with the given request body and headers.
     * @param url The URL of the SharePoint API endpoint.
     * @param body The body of the request, which can be an object (default is an empty object).
     * @param headers Additional headers to include in the request (default is an empty object).
     * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data or errors.
     */
    // REFACTOR: `return await` is redundant in all three delegate methods below (spDelete, spPatch, spPut).
    // `return this.spPost(...)` is sufficient and avoids an extra microtask tick.
    public async spDelete(url: string, body: object = {}, headers: object = {}): Promise<IApiResult> {
        return this.spPost(
            url,
            body,
            {
                ...headers,
                'X-Http-Method': 'DELETE',
                'IF-MATCH': '*'
            }
        );
    }

    /**
     * Sends a PATCH request (MERGE method) to a specified SharePoint API endpoint with the given request body and headers.
     * @param url The URL of the SharePoint API endpoint.
     * @param body The body of the request, which can be an object (default is an empty object).
     * @param headers Additional headers to include in the request (default is an empty object).
     * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data or errors.
     */
    public async spPatch(url: string, body: object | string = {}, headers: object = {}): Promise<IApiResult> {
        return this.spPost(
            url,
            body,
            {
                ...headers,
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE"
            }
        );
    }

    /**
     * Sends a PUT request to a specified SharePoint API endpoint with the given request body and headers.
     * @param url The URL of the SharePoint API endpoint.
     * @param body The body of the request, which can be an object or string (default is an empty object).
     * @param headers Additional headers to include in the request (default is an empty object).
     * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data or errors.
     */
    public async spPut(url: string, body: object | string = {}, headers: object = {}): Promise<IApiResult> {
        return this.spPost(
            url,
            body,
            {
                ...headers,
                "IF-MATCH": "*",
                "X-HTTP-Method": "PUT"
            }
        );
    }

    /**
     * Sends a GET request to a specified SharePoint API endpoint with the given headers.
     * @param url The URL of the SharePoint API endpoint.
     * @param headers Additional headers to include in the request (default is an empty object).
     * @returns A Promise that resolves to an IApiResult, containing the status of the request, any response data, or errors.
     */
    public async spGet(url: string, headers: object = {}, logError: boolean = true, responseType: "json" | "text" = "json"): Promise<IApiResult> {
        try {
            const spOpts: ISPHttpClientOptions = {
                headers: {
                    'Accept': 'application/json',
                    ...(responseType === "json" && { 'Content-type': 'application/json' }),
                    ...headers
                }
            };

            let results: any[] = [];
            let nextUrl: string | undefined = url;
            let lastResponseJson: any = null;
            let isPaginated = false;


            while (nextUrl) {
                const response: SPHttpClientResponse = await this.webPartContext.spHttpClient.get(nextUrl, SPHttpClient.configurations.v1, spOpts);
                let responseJson: any;
                if (response.ok && responseType === "text") {
                    const text = await response.text();
                    return {
                        ok: true,
                        code: 200,
                        data: text
                    };
                }

                try {
                    responseJson = await response.clone().json();
                } catch (_) {
                    if (!response.ok && logError) {
                        const error = await response.text();
                        console.error(error);
                    }
                }

                if (response.ok) {
                    lastResponseJson = responseJson;

                    if (responseJson.value) {
                        results = results.concat(responseJson.value);
                        isPaginated = true; // Means multiple pages exist
                    } else {
                        results.push(responseJson); // Single object case
                    }

                    nextUrl = responseJson["@odata.nextLink"]; // If exists, fetch next page
                } else {
                    if (logError) {
                        console.error(responseJson || `${response.status} ${response.statusText}`.trim());
                    }
                    return {
                        ok: false,
                        code: response.status,
                        exception: JSON.stringify(responseJson)
                    };
                }
            }

            return {
                ok: true,
                code: 200,
                data: isPaginated ? results : lastResponseJson // Keep original format if no paging
            };
        }
        catch (exception: any) {
            if (logError) {
                console.error('Error in GET request: ', url);
                console.error(exception.error, exception);
            }

            return {
                ok: false,
                code: 500,
                error: exception.message,
                exception
            };
        }
    }


    /**
     * Sends a GET request to retrieve binary data (as an ArrayBuffer) from a specified SharePoint API URL.
     * @param url The API endpoint to fetch the binary content from.
     * @param headers (Optional) Additional headers to include in the request.
     * @returns A Promise resolving to an ArrayBuffer containing the binary data, or throws an error if the request fails.
     */
    // REFACTOR: `spGetBinary` and `spPostBinary` break the uniform `IApiResult` return contract
    // used by all other methods. Consider returning `IApiResult<ArrayBuffer>` (generic) so callers
    // can handle failures consistently without try/catch at every call site.
    public async spGetBinary(url: string, headers: object = {}): Promise<IApiResult> {
        try {
            const spOpts: ISPHttpClientOptions = {
                headers: {
                    'Accept': 'application/octet-stream',
                    ...headers
                }
            };
            const response: SPHttpClientResponse = await this.webPartContext.spHttpClient.get(url, SPHttpClient.configurations.v1, spOpts);

            if (response.ok) {
                const arrayBuffer = await response.arrayBuffer();
                return {
                    ok: true,
                    data: arrayBuffer
                };
            } else {
                const error = await response.text();
                console.error('Error fetching binary data: ', error);
                return {
                    ok: false,
                    error
                };
            }
        }
        catch (exception: any) {
            console.error('Error in binary GET request: ', url);
            console.error(exception);
            return {
                ok: false,
                error: exception.message,
                exception
            };
        }
    }

    /**
     * Sends a POST request to upload binary data (as an ArrayBuffer) to a specified SharePoint API URL.
     * @param url The API endpoint to which the binary content should be uploaded.
     * @param fileContent The binary content to be uploaded, provided as an ArrayBuffer.
     * @returns A Promise that resolves when the upload is successful, or throws an error if the request fails.
     */
    public async spPostBinary(url: string, fileContent: ArrayBuffer): Promise<IApiResult> {
        try {
            const spOpts: ISPHttpClientOptions = {
                body: fileContent,  // Attach the binary content to the body
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/octet-stream', // Set the content type as binary
                }
            };

            const response: SPHttpClientResponse = await this.webPartContext.spHttpClient.post(url, SPHttpClient.configurations.v1, spOpts);

            if (!response.ok) {
                const error = await response.text();
                console.error('Error uploading binary data: ', error);
                throw new Error(error);
            }
            return {
                ok: true
            };
        } catch (exception) {
            console.error('Error in binary POST request: ', url);
            console.error(exception);
            return {
                ok: false
            };
        }
    }

    private async getGraphClient(url: string): Promise<any> {
        const GRAPH_API_VERSION = "3";
        const client = await this.webPartContext.msGraphClientFactory.getClient(GRAPH_API_VERSION);
        const graphClient = client.api(url);
        return graphClient;
    }

    /**
     * Sends a GET request to the Microsoft Graph API and retrieves data.
     * @param url The URL of the Microsoft Graph API endpoint.
     * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data.
     */
    // REFACTOR: `graphGet` and `graphPost` both repeat the same two-line client initialization.
    // Extract a private `getGraphClient(url: string)` helper to eliminate the duplication.
    // Also: the magic string "3" is the MSGraph API version — extract it to a named constant, e.g.:
    //   private static readonly GRAPH_API_VERSION = "3";
    public async graphGet(url: string, attempt: number = 1): Promise<IApiResult> {
        try {
            const graphClient = await this.getGraphClient(url);
            const response = await graphClient.get();
            if (!response) {
                return {
                    ok: false
                };
            }
            return {
                ok: true,
                data: response.value ?? response
            };

        } catch (exception: any) {
            if (exception?.statusCode === 429 && attempt < 3) {
                await delay(1000);
                return this.graphGet(url, attempt + 1);
            }
            return {
                ok: false
            };
        }
    }

    /**
     * Sends a POST request to the Microsoft Graph API with a specified body.
     * @param url The URL of the Microsoft Graph API endpoint.
     * @param body The body of the POST request, which is an object (default is an empty object).
     * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data.
     */
    public async graphPost(url: string, body: object = {}, attempt: number = 1): Promise<IApiResult> {
        try {
            const graphClient = await this.getGraphClient(url);
            const response = await graphClient.post(body);
            if (!response) {
                return {
                    ok: false
                };
            }
            return {
                ok: true,
                data: response.value ?? response,
            };
        } catch (exception: any) {
            if (exception?.statusCode === 429 && attempt < 3) {
                await delay(1000);
                return this.graphPost(url, body, attempt + 1);
            }
            return {
                ok: false
            };
        }
    }

    public async graphDelete(url: string, attempt: number = 1): Promise<IApiResult> {
        try {
            const graphClient = await this.getGraphClient(url);
            const response = await graphClient.delete();
            if (!response) {
                return {
                    ok: false
                };
            }
            return {
                ok: true,
                data: response.value ?? response,
            };
        } catch (exception: any) {
            if (exception?.statusCode === 429 && attempt < 3) {
                await delay(1000);
                return this.graphDelete(url, attempt + 1);
            }
            return {
                ok: false
            };
        }
    }
}