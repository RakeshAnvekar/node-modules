import { BaseComponentContext } from '@microsoft/sp-component-base';
import { IBrandFontToken } from '../propertyFields/brandFontPicker/IPropertyFieldBrandFontPicker';
/**
 * Service to interact with SharePoint Brand Center for font tokens
 */
export declare class BrandCenterService {
    private readonly context;
    constructor(context: BaseComponentContext);
    /**
     * Get font tokens from SharePoint Brand Center
     */
    getFontTokens(): Promise<IBrandFontToken[]>;
    /**
     * Get system font tokens as fallback
     */
    private getSystemFontTokens;
    /**
     * Try to get font tokens using SharePoint Brand Center REST API
     */
    private getFontTokensFromRest;
    /**
     * Fetch site font packages from SharePoint API
     */
    private fetchSiteFontPackages;
    /**
     * Parse the site font packages response
     */
    private parseSiteFontPackagesResponse;
    /**
     * Process JSON response containing font packages
     */
    private processJsonFontPackagesResponse;
    /**
     * Process Atom XML response containing font packages
     */
    private processAtomXmlFontPackagesResponse;
    /**
     * Parse an Atom entry element to extract font package data
     */
    private parseAtomEntry;
    /**
     * Process a site font package and extract font tokens
     */
    private processSiteFontPackage;
}
//# sourceMappingURL=BrandCenterService.d.ts.map