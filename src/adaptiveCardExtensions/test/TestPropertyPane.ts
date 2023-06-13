import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'TestAdaptiveCardExtensionStrings';
import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker';
import { PropertyFieldSpinner} from '@pnp/spfx-property-controls/lib/PropertyFieldSpinner';
import { ITestAdaptiveCardExtensionProps } from './TestAdaptiveCardExtension';
import { SpinnerSize } from '@fluentui/react/lib/Spinner';

export class TestPropertyPane {
  public getPropertyPaneConfiguration(properties: ITestAdaptiveCardExtensionProps, 
    onPropertyPaneFieldChanged: (propertyPath: string, oldValue: any, newValue: any) => void): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                //If you comment this two properties, everything works ok:
                PropertyFieldSpinner("", {
                  key: "sp1",
                  size: SpinnerSize.medium,
                  isVisible: true,
                  label: "Loading ..."
                }),
                PropertyFieldIconPicker('iconPicker', {
                  currentIcon: properties.iconPicker,
                  key: "iconPickerId",
                  onSave: (icon: string) => { console.log(icon); properties.iconPicker = icon; },
                  onChanged:(icon: string) => { console.log(icon);  },
                  buttonLabel: "Icon",
                  renderOption: "panel",
                  properties: properties,
                  onPropertyChange: onPropertyPaneFieldChanged,
                  label: "Icon Picker"              
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
