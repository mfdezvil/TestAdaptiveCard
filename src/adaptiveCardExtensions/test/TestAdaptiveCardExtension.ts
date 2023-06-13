import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { TestPropertyPane } from './TestPropertyPane';

export interface ITestAdaptiveCardExtensionProps {
  title: string;
  iconPicker: string;
}

export interface ITestAdaptiveCardExtensionState {
}

const CARD_VIEW_REGISTRY_ID: string = 'Test_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'Test_QUICK_VIEW';

export default class TestAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ITestAdaptiveCardExtensionProps,
  ITestAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: TestPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = { };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'Test-property-pane'*/
      './TestPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.TestPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration(this.properties, this.onPropertyPaneFieldChanged.bind(this));
  }
}
