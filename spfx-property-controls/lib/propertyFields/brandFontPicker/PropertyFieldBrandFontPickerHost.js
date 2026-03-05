var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
import * as React from 'react';
import { Dropdown, DropdownMenuItemType, Label, Spinner, SpinnerSize, MessageBar, MessageBarType } from '@fluentui/react';
import { BrandCenterService } from '../../services/BrandCenterService';
/**
 * Renders the controls for PropertyFieldBrandFontPicker component
 */
export default class PropertyFieldBrandFontPickerHost extends React.Component {
    constructor(props) {
        super(props);
        /**
         * Handle font selection change
         */
        this.onSelectionChanged = (event, option) => {
            if (option && option.itemType !== DropdownMenuItemType.Header) {
                const selectedToken = this.state.fontTokens.find(token => token.name === option.key);
                if (selectedToken) {
                    this.setState({ selectedToken });
                    if (this.props.onSelectionChanged) {
                        this.props.onSelectionChanged(selectedToken);
                    }
                }
            }
        };
        /**
         * Custom render for dropdown option to show font preview
         */
        this.onRenderOption = (option) => {
            if (!option) {
                return React.createElement("div", null);
            }
            // Skip rendering for header items
            if (option.itemType === DropdownMenuItemType.Header) {
                return (React.createElement("div", { style: {
                        fontWeight: '600',
                        color: '#605e5c',
                        padding: '8px 12px',
                        fontSize: '12px',
                        textTransform: 'uppercase',
                        letterSpacing: '0.5px'
                    } }, option.text));
            }
            const fontToken = this.state.fontTokens.find(token => token.name === option.key);
            const fontValue = (fontToken === null || fontToken === void 0 ? void 0 : fontToken.value) || '';
            return (React.createElement("div", { style: {
                    padding: '8px 12px',
                    minHeight: '40px',
                    display: 'flex',
                    flexDirection: 'column',
                    justifyContent: 'center'
                } },
                React.createElement("div", { style: {
                        fontSize: '14px',
                        color: '#323130',
                        lineHeight: '20px',
                        marginBottom: this.props.showPreview ? '2px' : '0'
                    } }, option.text),
                this.props.showPreview && (React.createElement("div", { style: {
                        fontFamily: fontValue,
                        fontSize: '12px',
                        color: '#605e5c',
                        lineHeight: '16px'
                    } }, "Sample text preview"))));
        };
        /**
         * Custom render for dropdown title to show selected font preview
         */
        this.onRenderTitle = (options) => {
            if (!options || options.length === 0) {
                return React.createElement("div", null);
            }
            const option = options[0];
            const fontToken = this.state.fontTokens.find(token => token.name === option.key);
            const fontValue = (fontToken === null || fontToken === void 0 ? void 0 : fontToken.value) || '';
            return (React.createElement("div", { style: {
                    display: 'flex',
                    flexDirection: 'column',
                    justifyContent: 'center',
                    height: '40px',
                    paddingLeft: '12px',
                    paddingRight: '12px'
                } },
                React.createElement("span", { style: {
                        color: '#323130',
                        fontSize: '14px',
                        lineHeight: '20px'
                    } }, option.text),
                this.props.showPreview && (React.createElement("span", { style: {
                        fontFamily: fontValue,
                        fontSize: '12px',
                        color: '#605e5c',
                        lineHeight: '16px'
                    } }, "Sample text preview"))));
        };
        this.state = {
            loading: true,
            fontTokens: [],
            selectedToken: undefined,
            errorMessage: undefined
        };
        // Initialize the BrandCenterService
        this.brandCenterService = new BrandCenterService(this.props.context);
    }
    componentDidMount() {
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        this.loadFontTokens();
    }
    /**
     * Load font tokens from Brand Center or fallback
     */
    loadFontTokens() {
        return __awaiter(this, void 0, void 0, function* () {
            this.setState({ loading: true, errorMessage: undefined });
            try {
                let fontTokens = [];
                console.log('🎨 Brand Center Font Picker: Starting font token loading...');
                // Check if custom font tokens are provided
                if (this.props.customFontTokens && this.props.customFontTokens.length > 0) {
                    fontTokens = this.props.customFontTokens;
                    console.log('🎨 Brand Center Font Picker: Using custom font tokens:', fontTokens.length);
                }
                else {
                    // Try to load from Brand Center using the service
                    console.log('🎨 Brand Center Font Picker: Loading from Brand Center service...');
                    fontTokens = yield this.brandCenterService.getFontTokens();
                    console.log('🎨 Brand Center Font Picker: Loaded font tokens:', fontTokens.length, fontTokens);
                }
                // Set initial selected token
                let selectedToken;
                if (this.props.initialValue) {
                    selectedToken = fontTokens.find(token => token.value === this.props.initialValue);
                }
                this.setState({
                    loading: false,
                    fontTokens,
                    selectedToken,
                    errorMessage: undefined
                });
                // Notify parent component
                if (this.props.onFontTokensLoaded) {
                    this.props.onFontTokensLoaded(fontTokens);
                }
            }
            catch (error) {
                console.error('Error loading font tokens:', error);
                const errorMessage = this.props.loadingErrorMessage || 'Failed to load font tokens';
                this.setState({
                    loading: false,
                    fontTokens: [],
                    selectedToken: undefined,
                    errorMessage
                });
            }
        });
    }
    render() {
        const { label, disabled } = this.props;
        const { loading, fontTokens, selectedToken, errorMessage } = this.state;
        // Group font tokens by category and convert to dropdown options
        const options = [];
        // Group fonts by category
        const categorizedFonts = {
            site: fontTokens.filter(token => token.category === 'site'),
            microsoft: fontTokens.filter(token => token.category === 'microsoft')
        };
        // Add "From this site" section
        if (categorizedFonts.site.length > 0) {
            options.push({
                key: 'site-header',
                text: 'From this site',
                itemType: DropdownMenuItemType.Header
            });
            categorizedFonts.site.forEach(token => {
                options.push({
                    key: token.name,
                    text: token.displayName
                });
            });
        }
        // Add "From Microsoft" section
        if (categorizedFonts.microsoft.length > 0) {
            options.push({
                key: 'microsoft-header',
                text: 'From Microsoft',
                itemType: DropdownMenuItemType.Header
            });
            categorizedFonts.microsoft.forEach(token => {
                options.push({
                    key: token.name,
                    text: token.displayName
                });
            });
        }
        const selectedKey = selectedToken ? selectedToken.name : undefined;
        const dropdownStyles = {
            dropdown: {
                width: '100%'
            },
            title: {
                height: '40px',
                lineHeight: '40px'
            },
            callout: {
                maxHeight: '300px'
            }
        };
        return (React.createElement("div", null,
            label && React.createElement(Label, null, label),
            loading && (React.createElement("div", { style: { padding: '10px 0' } },
                React.createElement(Spinner, { size: SpinnerSize.small, label: "Loading font tokens..." }))),
            errorMessage && (React.createElement(MessageBar, { messageBarType: MessageBarType.error, style: { marginBottom: '10px' } }, errorMessage)),
            !loading && !errorMessage && (React.createElement(Dropdown, { options: options, selectedKey: selectedKey, onChange: this.onSelectionChanged, disabled: disabled, placeholder: "Select a font...", styles: dropdownStyles, onRenderOption: this.props.showPreview ? this.onRenderOption : undefined, onRenderTitle: this.props.showPreview ? this.onRenderTitle : undefined }))));
    }
}
//# sourceMappingURL=PropertyFieldBrandFontPickerHost.js.map