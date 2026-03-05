var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
/**
 * Service to interact with SharePoint Brand Center for font tokens
 */
export class BrandCenterService {
    constructor(context) {
        this.context = context;
    }
    /**
     * Get font tokens from SharePoint Brand Center
     */
    getFontTokens() {
        return __awaiter(this, void 0, void 0, function* () {
            const siteFontTokens = yield this.getFontTokensFromRest();
            const systemTokens = this.getSystemFontTokens();
            // Combine all font tokens with categories
            const allTokens = [...siteFontTokens, ...systemTokens];
            return allTokens;
        });
    }
    /**
     * Get system font tokens as fallback
     */
    getSystemFontTokens() {
        return [
            {
                name: 'fontFamilyBase',
                displayName: 'Base Font',
                value: 'var(--fontFamilyBase, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif)',
                preview: 'The quick brown fox jumps over the lazy dog',
                category: 'microsoft'
            },
            {
                name: 'fontFamilyMonospace',
                displayName: 'Monospace Font',
                value: 'var(--fontFamilyMonospace, Consolas, "Courier New", Courier, monospace)',
                preview: 'The quick brown fox jumps over the lazy dog',
                category: 'microsoft'
            },
            {
                name: 'fontFamilyNumeric',
                displayName: 'Numeric Font',
                value: 'var(--fontFamilyNumeric, Bahnschrift, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif)',
                preview: '0123456789',
                category: 'microsoft'
            }
        ];
    }
    /**
     * Try to get font tokens using SharePoint Brand Center REST API
     */
    getFontTokensFromRest() {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const spHttpClient = this.context.spHttpClient;
                const currentWebUrl = this.context.pageContext.web.absoluteUrl;
                return yield this.fetchSiteFontPackages(spHttpClient, currentWebUrl);
            }
            catch (error) {
                console.debug('SharePoint Brand Center REST API access not available:', error);
            }
            return [];
        });
    }
    /**
     * Fetch site font packages from SharePoint API
     */
    fetchSiteFontPackages(spHttpClient, currentWebUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const tokens = [];
            try {
                const siteFontPackagesResponse = yield spHttpClient.get(`${currentWebUrl}/_api/SiteFontPackages`, {
                    headers: {
                        'Accept': 'application/json;odata=verbose',
                        'Content-Type': 'application/json;odata=verbose'
                    }
                } // eslint-disable-line @typescript-eslint/no-explicit-any
                );
                if (siteFontPackagesResponse.ok) {
                    return yield this.parseSiteFontPackagesResponse(siteFontPackagesResponse, currentWebUrl);
                }
                else {
                    console.debug(`Site font packages API returned ${siteFontPackagesResponse.status}: ${siteFontPackagesResponse.statusText}`);
                }
            }
            catch (siteFontPackagesError) {
                console.debug('Site font packages not available:', siteFontPackagesError);
            }
            return tokens;
        });
    }
    /**
     * Parse the site font packages response
     */
    parseSiteFontPackagesResponse(response, currentWebUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const contentType = response.headers.get('content-type');
            // Check if the response is JSON
            if (contentType === null || contentType === void 0 ? void 0 : contentType.includes('application/json')) {
                return yield this.processJsonFontPackagesResponse(response, currentWebUrl);
            }
            // Check if the response is Atom XML feed
            else if (contentType === null || contentType === void 0 ? void 0 : contentType.includes('application/atom+xml')) {
                return yield this.processAtomXmlFontPackagesResponse(response, currentWebUrl);
            }
            else {
                // Response is not JSON or Atom XML, likely an error response
                const responseText = yield response.text();
                console.debug('Site font packages returned unexpected response:', responseText.substring(0, 200));
                return [];
            }
        });
    }
    /**
     * Process JSON response containing font packages
     */
    processJsonFontPackagesResponse(response, currentWebUrl) {
        var _a;
        return __awaiter(this, void 0, void 0, function* () {
            const tokens = [];
            const siteFontPackagesData = yield response.json();
            if ((_a = siteFontPackagesData === null || siteFontPackagesData === void 0 ? void 0 : siteFontPackagesData.d) === null || _a === void 0 ? void 0 : _a.results) {
                for (const fontPackage of siteFontPackagesData.d.results) {
                    if (fontPackage.ID && !fontPackage.IsHidden && fontPackage.IsValid) {
                        const fontTokens = yield this.processSiteFontPackage(fontPackage, currentWebUrl);
                        if (fontTokens.length > 0) {
                            tokens.push(...fontTokens);
                        }
                    }
                }
            }
            return tokens;
        });
    }
    /**
     * Process Atom XML response containing font packages
     */
    processAtomXmlFontPackagesResponse(response, currentWebUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const tokens = [];
            try {
                const xmlText = yield response.text();
                // Parse the Atom XML feed - look for entry elements
                const parser = new DOMParser();
                const xmlDoc = parser.parseFromString(xmlText, 'text/xml');
                // Check for parsing errors
                const parserErrors = xmlDoc.getElementsByTagName('parsererror');
                if (parserErrors.length > 0) {
                    console.debug('XML parsing error:', parserErrors[0].textContent);
                    return tokens;
                }
                // Look for entry elements in the Atom feed
                const entries = xmlDoc.getElementsByTagName('entry');
                for (const entry of Array.from(entries)) {
                    const fontPackage = this.parseAtomEntry(entry);
                    if ((fontPackage === null || fontPackage === void 0 ? void 0 : fontPackage.ID) && !fontPackage.IsHidden && fontPackage.IsValid) {
                        const fontTokens = yield this.processSiteFontPackage(fontPackage, currentWebUrl);
                        if (fontTokens.length > 0) {
                            tokens.push(...fontTokens);
                        }
                    }
                }
            }
            catch (xmlParseError) {
                console.debug('Could not parse Atom XML response:', xmlParseError);
            }
            return tokens;
        });
    }
    /**
     * Parse an Atom entry element to extract font package data
     */
    parseAtomEntry(entry) {
        try {
            const properties = entry.getElementsByTagName('m:properties')[0];
            if (!properties) {
                return null;
            }
            const getValue = (tagName) => {
                const element = properties.getElementsByTagName(tagName)[0];
                return (element === null || element === void 0 ? void 0 : element.textContent) || null;
            };
            const getBoolValue = (tagName) => {
                const value = getValue(tagName);
                return value === 'true';
            };
            return {
                ID: getValue('d:ID'),
                Title: getValue('d:Title'),
                IsHidden: getBoolValue('d:IsHidden'),
                IsValid: getBoolValue('d:IsValid'),
                PackageJson: getValue('d:PackageJson')
            };
        }
        catch (parseError) {
            console.debug('Could not parse Atom entry:', parseError);
            return null;
        }
    }
    /**
     * Process a site font package and extract font tokens
     */
    processSiteFontPackage(fontPackage, webUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const tokens = [];
            try {
                // Parse the PackageJson to get font information
                if (fontPackage.PackageJson) {
                    const packageData = JSON.parse(fontPackage.PackageJson);
                    // Create a single token for the entire font package
                    const packageTitle = fontPackage.Title; // Use the title as is
                    // Determine the primary font value to use for the package
                    let primaryFontValue = `"${packageTitle}", sans-serif`;
                    // Try to get the primary font from font slots or faces
                    if (packageData.fontSlots) {
                        // Prefer body font, then heading, then title, then label
                        const preferenceOrder = ['body', 'heading', 'title', 'label'];
                        for (const slotName of preferenceOrder) {
                            const slot = packageData.fontSlots[slotName];
                            if (slot === null || slot === void 0 ? void 0 : slot.fontFamily) {
                                primaryFontValue = `"${slot.fontFamily}", sans-serif`;
                                break;
                            }
                        }
                    }
                    else if (packageData.fontFaces && Array.isArray(packageData.fontFaces) && packageData.fontFaces.length > 0) {
                        // Use the first font face if no slots are available
                        const firstFontFace = packageData.fontFaces[0];
                        if (firstFontFace === null || firstFontFace === void 0 ? void 0 : firstFontFace.fontFamily) {
                            primaryFontValue = `"${firstFontFace.fontFamily}", sans-serif`;
                        }
                    }
                    // Create a single token for this font package
                    tokens.push({
                        name: `siteFontPackage${fontPackage.ID.replace(/[^a-zA-Z0-9]/g, '')}`,
                        displayName: packageTitle,
                        value: primaryFontValue,
                        preview: 'The quick brown fox jumps over the lazy dog',
                        fileUrl: `${webUrl}/_api/SiteFontPackages/GetById('${fontPackage.ID}')`,
                        category: 'site'
                    });
                    // If no specific fonts found, create a general token from the package title
                    if (tokens.length === 0 && fontPackage.Title) {
                        tokens.push({
                            name: `siteFontPackage${fontPackage.ID.replace(/[^a-zA-Z0-9]/g, '')}`,
                            displayName: `Brand Font: ${packageTitle}`,
                            value: `"${packageTitle}", sans-serif`,
                            preview: 'The quick brown fox jumps over the lazy dog',
                            fileUrl: `${webUrl}/_api/SiteFontPackages/GetById('${fontPackage.ID}')`,
                            category: 'site'
                        });
                    }
                }
            }
            catch (parseError) {
                console.debug(`Could not parse font package ${fontPackage.ID}:`, parseError);
                // Fallback: create token from title if JSON parsing fails
                if (fontPackage.Title) {
                    const packageTitle = fontPackage.Title; // Use the title as is
                    tokens.push({
                        name: `siteFontPackage${fontPackage.ID.replace(/[^a-zA-Z0-9]/g, '')}`,
                        displayName: `Brand Font: ${packageTitle}`,
                        value: `"${packageTitle}", sans-serif`,
                        preview: 'The quick brown fox jumps over the lazy dog',
                        fileUrl: `${webUrl}/_api/SiteFontPackages/GetById('${fontPackage.ID}')`,
                        category: 'site'
                    });
                }
            }
            return tokens;
        });
    }
}
//# sourceMappingURL=BrandCenterService.js.map