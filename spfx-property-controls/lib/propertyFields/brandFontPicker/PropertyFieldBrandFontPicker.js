import * as React from 'react';
import * as ReactDom from 'react-dom';
import { PropertyPaneFieldType } from '@microsoft/sp-property-pane';
import { setPropertyValue } from '../../helpers/GeneralHelper';
import { debounce } from '../../common/util/Util';
import PropertyFieldBrandFontPickerHost from './PropertyFieldBrandFontPickerHost';
/**
 * Represents a PropertyFieldBrandFontPicker object
 */
class PropertyFieldBrandFontPickerBuilder {
    constructor(_targetProperty, _properties) {
        // Properties defined by IPropertyPaneField
        this.type = PropertyPaneFieldType.Custom;
        this._debounce = debounce(); // eslint-disable-line @typescript-eslint/no-explicit-any
        this.targetProperty = _targetProperty;
        this.properties = {
            key: _properties.key,
            label: _properties.label,
            targetProperty: _targetProperty,
            onPropertyChange: _properties.onPropertyChange,
            initialValue: _properties.initialValue,
            disabled: _properties.disabled,
            isHidden: _properties.isHidden,
            context: _properties.context,
            properties: _properties.properties,
            customFontTokens: _properties.customFontTokens,
            onFontTokensLoaded: _properties.onFontTokensLoaded,
            showPreview: _properties.showPreview,
            previewText: _properties.previewText,
            loadingErrorMessage: _properties.loadingErrorMessage,
            useSystemFallback: _properties.useSystemFallback,
            onRender: this.onRender.bind(this),
            onDispose: this.onDispose.bind(this)
        };
    }
    /**
     * Renders the BrandFontPicker field content
     */
    onRender(elem, ctx, changeCallback) {
        if (!this.elem) {
            this.elem = elem;
        }
        this.changeCB = changeCallback;
        const element = React.createElement(PropertyFieldBrandFontPickerHost, {
            label: this.properties.label,
            targetProperty: this.targetProperty,
            context: this.properties.context,
            initialValue: this.properties.initialValue,
            disabled: this.properties.disabled,
            onSelectionChanged: this.onSelectionChanged.bind(this),
            customFontTokens: this.properties.customFontTokens,
            onFontTokensLoaded: this.properties.onFontTokensLoaded,
            showPreview: this.properties.showPreview,
            previewText: this.properties.previewText,
            loadingErrorMessage: this.properties.loadingErrorMessage,
            useSystemFallback: this.properties.useSystemFallback
        });
        ReactDom.render(element, elem);
    }
    /**
     * Disposes the current object
     */
    onDispose(elem) {
        ReactDom.unmountComponentAtNode(elem);
    }
    /**
     * Called when the font selection has been changed
     */
    onSelectionChanged(option) {
        this._debounce(() => {
            setPropertyValue(this.properties.properties, this.targetProperty, option.value);
            this.properties.onPropertyChange(this.targetProperty, this.properties.initialValue, option.value);
            if (typeof this.changeCB !== 'undefined' && this.changeCB !== null) {
                this.changeCB(this.targetProperty, option.value);
            }
        }, 200);
    }
}
/**
 * Helper method to create a Brand Font Picker on the PropertyPane.
 * @param targetProperty - Target property the Brand Font Picker is associated to.
 * @param properties - Strongly typed Brand Font Picker properties.
 */
export function PropertyFieldBrandFontPicker(targetProperty, properties) {
    // Create an internal properties object from the given properties
    const newProperties = {
        label: properties.label,
        onPropertyChange: properties.onPropertyChange,
        context: properties.context,
        initialValue: properties.initialValue,
        disabled: properties.disabled,
        isHidden: properties.isHidden,
        properties: properties.properties,
        customFontTokens: properties.customFontTokens,
        onFontTokensLoaded: properties.onFontTokensLoaded,
        showPreview: properties.showPreview !== false, // Default to true
        previewText: properties.previewText,
        loadingErrorMessage: properties.loadingErrorMessage,
        useSystemFallback: properties.useSystemFallback !== false, // Default to true
        key: properties.key
    };
    // Safely set the property
    setPropertyValue(newProperties.properties, targetProperty, newProperties.initialValue);
    // Create a new instance of the PropertyFieldBrandFontPickerBuilder
    return new PropertyFieldBrandFontPickerBuilder(targetProperty, newProperties);
}
//# sourceMappingURL=PropertyFieldBrandFontPicker.js.map