import * as React from 'react';
import { IPropertyFieldBrandFontPickerHostProps, IPropertyFieldBrandFontPickerHostState } from './IPropertyFieldBrandFontPickerHost';
/**
 * Renders the controls for PropertyFieldBrandFontPicker component
 */
export default class PropertyFieldBrandFontPickerHost extends React.Component<IPropertyFieldBrandFontPickerHostProps, IPropertyFieldBrandFontPickerHostState> {
    private readonly brandCenterService;
    constructor(props: IPropertyFieldBrandFontPickerHostProps);
    componentDidMount(): void;
    /**
     * Load font tokens from Brand Center or fallback
     */
    private loadFontTokens;
    /**
     * Handle font selection change
     */
    private readonly onSelectionChanged;
    /**
     * Custom render for dropdown option to show font preview
     */
    private readonly onRenderOption;
    /**
     * Custom render for dropdown title to show selected font preview
     */
    private readonly onRenderTitle;
    render(): React.ReactElement<IPropertyFieldBrandFontPickerHostProps>;
}
//# sourceMappingURL=PropertyFieldBrandFontPickerHost.d.ts.map