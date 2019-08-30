// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
#include "pch.h"

#include "AdaptiveActionSet.h"
#include "AdaptiveActionSetRenderer.h"
#include "AdaptiveShowCardActionRenderer.h"
#include "AdaptiveElementParserRegistration.h"
#include "AdaptiveImage.h"
#include "AdaptiveRenderArgs.h"
#include "ActionHelpers.h"

using namespace Microsoft::WRL;
using namespace Microsoft::WRL::Wrappers;
using namespace ABI::AdaptiveNamespace;
using namespace ABI::Windows::Foundation;
using namespace ABI::Windows::Foundation::Collections;
using namespace ABI::Windows::UI::Xaml;
using namespace ABI::Windows::UI::Xaml::Controls;
using namespace ABI::Windows::UI::Xaml::Controls::Primitives;

namespace AdaptiveNamespace
{
    HRESULT AdaptiveActionSetRenderer::RuntimeClassInitialize() noexcept try
    {
        return S_OK;
    }
    CATCH_RETURN;

    HRESULT AdaptiveActionSetRenderer::FromJson(
        _In_ ABI::Windows::Data::Json::IJsonObject* jsonObject,
        _In_ ABI::AdaptiveNamespace::IAdaptiveElementParserRegistration* elementParserRegistration,
        _In_ ABI::AdaptiveNamespace::IAdaptiveActionParserRegistration* actionParserRegistration,
        _In_ ABI::Windows::Foundation::Collections::IVector<ABI::AdaptiveNamespace::AdaptiveWarning*>* adaptiveWarnings,
        _COM_Outptr_ ABI::AdaptiveNamespace::IAdaptiveCardElement** element) noexcept try
    {
        return AdaptiveNamespace::FromJson<AdaptiveNamespace::AdaptiveActionSet, AdaptiveSharedNamespace::ActionSet, AdaptiveSharedNamespace::ActionSetParser>(
            jsonObject, elementParserRegistration, actionParserRegistration, adaptiveWarnings, element);
    }
    CATCH_RETURN;

    template<typename T>
    static HRESULT TryGetResourceFromResourceDictionaries(_In_ IResourceDictionary* resourceDictionary,
                                                          std::wstring resourceName,
                                                          _COM_Outptr_ T** style) noexcept
    {
        if (resourceDictionary == nullptr)
        {
            return E_INVALIDARG;
        }

        *style = nullptr;
        try
        {
            // Get a resource key for the requested style that we can use for ResourceDictionary Lookups
            ComPtr<IPropertyValueStatics> propertyValueStatics;
            THROW_IF_FAILED(GetActivationFactory(HStringReference(RuntimeClass_Windows_Foundation_PropertyValue).Get(),
                                                 &propertyValueStatics));
            ComPtr<IInspectable> resourceKey;
            THROW_IF_FAILED(propertyValueStatics->CreateString(HStringReference(resourceName.c_str()).Get(),
                                                               resourceKey.GetAddressOf()));

            // Search for the named resource
            ComPtr<IResourceDictionary> strongDictionary = resourceDictionary;
            ComPtr<IInspectable> dictionaryValue;
            ComPtr<IMap<IInspectable*, IInspectable*>> resourceDictionaryMap;

            boolean hasKey{};
            if (SUCCEEDED(strongDictionary.As(&resourceDictionaryMap)) &&
                SUCCEEDED(resourceDictionaryMap->HasKey(resourceKey.Get(), &hasKey)) && hasKey &&
                SUCCEEDED(resourceDictionaryMap->Lookup(resourceKey.Get(), dictionaryValue.GetAddressOf())))
            {
                ComPtr<T> resourceToReturn;
                if (SUCCEEDED(dictionaryValue.As(&resourceToReturn)))
                {
                    THROW_IF_FAILED(resourceToReturn.CopyTo(style));
                    return S_OK;
                }
            }
        }
        catch (...)
        {
        }
        return E_FAIL;
    }

    static HRESULT BuildActionSetHelper(_In_opt_ ABI::AdaptiveNamespace::IAdaptiveCard* adaptiveCard,
                                        _In_opt_ IAdaptiveActionSet* adaptiveActionSet,
                                        _In_ IVector<IAdaptiveActionElement*>* children,
                                        _In_ IAdaptiveRenderContext* renderContext,
                                        _In_ IAdaptiveRenderArgs* renderArgs,
                                        _Outptr_ IUIElement** actionSetControl)
    {
        ComPtr<IAdaptiveHostConfig> hostConfig;
        RETURN_IF_FAILED(renderContext->get_HostConfig(&hostConfig));
        ComPtr<IAdaptiveActionsConfig> actionsConfig;
        RETURN_IF_FAILED(hostConfig->get_Actions(actionsConfig.GetAddressOf()));

        ABI::AdaptiveNamespace::ActionAlignment actionAlignment;
        RETURN_IF_FAILED(actionsConfig->get_ActionAlignment(&actionAlignment));

        ABI::AdaptiveNamespace::ActionsOrientation actionsOrientation;
        RETURN_IF_FAILED(actionsConfig->get_ActionsOrientation(&actionsOrientation));

        // Declare the panel that will host the buttons
        ComPtr<IPanel> actionsPanel;
        ComPtr<IVector<ColumnDefinition*>> columnDefinitions;

        if (actionAlignment == ABI::AdaptiveNamespace::ActionAlignment::Stretch &&
            actionsOrientation == ABI::AdaptiveNamespace::ActionsOrientation::Horizontal)
        {
            // If stretch alignment and orientation is horizontal, we use a grid with equal column widths to achieve
            // stretch behavior. For vertical orientation, we'll still just use a stack panel since the concept of
            // stretching buttons height isn't really valid, especially when the height of cards are typically dynamic.
            ComPtr<IGrid> actionsGrid =
                XamlHelpers::CreateXamlClass<IGrid>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_Grid));
            RETURN_IF_FAILED(actionsGrid->get_ColumnDefinitions(&columnDefinitions));
            RETURN_IF_FAILED(actionsGrid.As(&actionsPanel));
        }
        else
        {
            // Create a stack panel for the action buttons
            ComPtr<IStackPanel> actionStackPanel =
                XamlHelpers::CreateXamlClass<IStackPanel>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_StackPanel));

            auto uiOrientation = (actionsOrientation == ABI::AdaptiveNamespace::ActionsOrientation::Horizontal) ?
                Orientation::Orientation_Horizontal :
                Orientation::Orientation_Vertical;

            RETURN_IF_FAILED(actionStackPanel->put_Orientation(uiOrientation));

            ComPtr<IFrameworkElement> actionsFrameworkElement;
            RETURN_IF_FAILED(actionStackPanel.As(&actionsFrameworkElement));

            switch (actionAlignment)
            {
            case ABI::AdaptiveNamespace::ActionAlignment::Center:
                RETURN_IF_FAILED(actionsFrameworkElement->put_HorizontalAlignment(ABI::Windows::UI::Xaml::HorizontalAlignment_Center));
                break;
            case ABI::AdaptiveNamespace::ActionAlignment::Left:
                RETURN_IF_FAILED(actionsFrameworkElement->put_HorizontalAlignment(ABI::Windows::UI::Xaml::HorizontalAlignment_Left));
                break;
            case ABI::AdaptiveNamespace::ActionAlignment::Right:
                RETURN_IF_FAILED(actionsFrameworkElement->put_HorizontalAlignment(ABI::Windows::UI::Xaml::HorizontalAlignment_Right));
                break;
            case ABI::AdaptiveNamespace::ActionAlignment::Stretch:
                RETURN_IF_FAILED(actionsFrameworkElement->put_HorizontalAlignment(ABI::Windows::UI::Xaml::HorizontalAlignment_Stretch));
                break;
            }

            // Add the action buttons to the stack panel
            RETURN_IF_FAILED(actionStackPanel.As(&actionsPanel));
        }

        Thickness buttonMargin;
        RETURN_IF_FAILED(ActionHelpers::GetButtonMargin(actionsConfig.Get(), buttonMargin));
        if (actionsOrientation == ABI::AdaptiveNamespace::ActionsOrientation::Horizontal)
        {
            // Negate the spacing on the sides so the left and right buttons are flush on the side.
            // We do NOT remove the margin from the individual button itself, since that would cause
            // the equal columns stretch behavior to not have equal columns (since the first and last
            // button would be narrower without the same margins as its peers).
            ComPtr<IFrameworkElement> actionsPanelAsFrameworkElement;
            RETURN_IF_FAILED(actionsPanel.As(&actionsPanelAsFrameworkElement));
            RETURN_IF_FAILED(actionsPanelAsFrameworkElement->put_Margin({buttonMargin.Left * -1, 0, buttonMargin.Right * -1, 0}));
        }
        else
        {
            // Negate the spacing on the top and bottom so the first and last buttons don't have extra padding
            ComPtr<IFrameworkElement> actionsPanelAsFrameworkElement;
            RETURN_IF_FAILED(actionsPanel.As(&actionsPanelAsFrameworkElement));
            RETURN_IF_FAILED(actionsPanelAsFrameworkElement->put_Margin({0, buttonMargin.Top * -1, 0, buttonMargin.Bottom * -1}));
        }

        UINT32 maxActions;
        RETURN_IF_FAILED(actionsConfig->get_MaxActions(&maxActions));

        bool allActionsHaveIcons{true};
        XamlHelpers::IterateOverVector<IAdaptiveActionElement>(children, [&](IAdaptiveActionElement* child) {
            HSTRING iconUrl;
            RETURN_IF_FAILED(child->get_IconUrl(&iconUrl));

            bool iconUrlIsEmpty = WindowsIsStringEmpty(iconUrl);
            if (iconUrlIsEmpty)
            {
                allActionsHaveIcons = false;
            }
            return S_OK;
        });

        UINT currentAction = 0;

        RETURN_IF_FAILED(renderArgs->put_AllowAboveTitleIconPlacement(allActionsHaveIcons));

        std::shared_ptr<std::vector<ComPtr<IUIElement>>> allShowCards = std::make_shared<std::vector<ComPtr<IUIElement>>>();
        ComPtr<IStackPanel> showCardsStackPanel =
            XamlHelpers::CreateXamlClass<IStackPanel>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_StackPanel));
        ComPtr<IGridStatics> gridStatics;
        RETURN_IF_FAILED(GetActivationFactory(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_Grid).Get(), &gridStatics));
        XamlHelpers::IterateOverVector<IAdaptiveActionElement>(children, [&](IAdaptiveActionElement* child) {
            if (currentAction < maxActions)
            {
                // Render each action using the registered renderer
                ComPtr<IAdaptiveActionElement> action(child);
                ComPtr<IAdaptiveActionRendererRegistration> actionRegistration;
                RETURN_IF_FAILED(renderContext->get_ActionRenderers(&actionRegistration));

                ComPtr<IAdaptiveActionRenderer> renderer;
                while (!renderer)
                {
                    HString actionTypeString;
                    RETURN_IF_FAILED(action->get_ActionTypeString(actionTypeString.GetAddressOf()));
                    RETURN_IF_FAILED(actionRegistration->Get(actionTypeString.Get(), &renderer));
                    if (!renderer)
                    {
                        ABI::AdaptiveNamespace::FallbackType actionFallbackType;
                        action->get_FallbackType(&actionFallbackType);
                        switch (actionFallbackType)
                        {
                        case ABI::AdaptiveNamespace::FallbackType::Drop:
                        {
                            RETURN_IF_FAILED(XamlBuilder::WarnForFallbackDrop(renderContext, actionTypeString.Get()));
                            return S_OK;
                        }

                        case ABI::AdaptiveNamespace::FallbackType::Content:
                        {
                            ComPtr<IAdaptiveActionElement> actionFallback;
                            RETURN_IF_FAILED(action->get_FallbackContent(&actionFallback));

                            HString fallbackTypeString;
                            RETURN_IF_FAILED(actionFallback->get_ActionTypeString(fallbackTypeString.GetAddressOf()));
                            RETURN_IF_FAILED(
                                XamlBuilder::WarnForFallbackContentElement(renderContext, actionTypeString.Get(), fallbackTypeString.Get()));

                            action = actionFallback;
                            break;
                        }

                        case ABI::AdaptiveNamespace::FallbackType::None:
                        default:
                            return E_FAIL;
                        }
                    }
                }

                ComPtr<IUIElement> actionControl;
                RETURN_IF_FAILED(renderer->Render(action.Get(), renderContext, renderArgs, &actionControl));

                XamlHelpers::AppendXamlElementToPanel(actionControl.Get(), actionsPanel.Get());

                ABI::AdaptiveNamespace::ActionType actionType;
                RETURN_IF_FAILED(action->get_ActionType(&actionType));

                // Build inline show cards if needed
                if (actionType == ABI::AdaptiveNamespace::ActionType_ShowCard)
                {
                    ComPtr<IUIElement> uiShowCard;

                    ComPtr<IAdaptiveShowCardActionConfig> showCardActionConfig;
                    RETURN_IF_FAILED(actionsConfig->get_ShowCard(&showCardActionConfig));

                    ABI::AdaptiveNamespace::ActionMode showCardActionMode;
                    RETURN_IF_FAILED(showCardActionConfig->get_ActionMode(&showCardActionMode));

                    if (showCardActionMode == ABI::AdaptiveNamespace::ActionMode::Inline)
                    {
                        ComPtr<IAdaptiveShowCardAction> showCardAction;
                        RETURN_IF_FAILED(action.As(&showCardAction));

                        ComPtr<IAdaptiveCard> showCard;
                        RETURN_IF_FAILED(showCardAction->get_Card(&showCard));

                        RETURN_IF_FAILED(AdaptiveShowCardActionRenderer::BuildShowCard(
                            showCard.Get(), renderContext, renderArgs, (adaptiveActionSet == nullptr), uiShowCard.GetAddressOf()));

                        ComPtr<IPanel> showCardsPanel;
                        RETURN_IF_FAILED(showCardsStackPanel.As(&showCardsPanel));
                        XamlHelpers::AppendXamlElementToPanel(uiShowCard.Get(), showCardsPanel.Get());

                        if (adaptiveActionSet)
                        {
                            RETURN_IF_FAILED(
                                renderContext->AddInlineShowCard(adaptiveActionSet, showCardAction.Get(), uiShowCard.Get()));
                        }
                        else
                        {
                            ComPtr<AdaptiveNamespace::AdaptiveRenderContext> contextImpl =
                                PeekInnards<AdaptiveNamespace::AdaptiveRenderContext>(renderContext);

                            RETURN_IF_FAILED(
                                contextImpl->AddInlineShowCard(adaptiveCard, showCardAction.Get(), uiShowCard.Get()));
                        }
                    }
                }

                if (columnDefinitions != nullptr)
                {
                    // If using the equal width columns, we'll add a column and assign the column
                    ComPtr<IColumnDefinition> columnDefinition = XamlHelpers::CreateXamlClass<IColumnDefinition>(
                        HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_ColumnDefinition));
                    RETURN_IF_FAILED(columnDefinition->put_Width({1.0, GridUnitType::GridUnitType_Star}));
                    RETURN_IF_FAILED(columnDefinitions->Append(columnDefinition.Get()));

                    ComPtr<IFrameworkElement> actionFrameworkElement;
                    THROW_IF_FAILED(actionControl.As(&actionFrameworkElement));
                    THROW_IF_FAILED(gridStatics->SetColumn(actionFrameworkElement.Get(), currentAction));
                }
            }
            else
            {
                renderContext->AddWarning(ABI::AdaptiveNamespace::WarningStatusCode::MaxActionsExceeded,
                                          HStringReference(L"Some actions were not rendered due to exceeding the maximum number of actions allowed")
                                              .Get());
            }
            currentAction++;
            return S_OK;
        });

        // Reset icon placement value
        RETURN_IF_FAILED(renderArgs->put_AllowAboveTitleIconPlacement(false));

        ComPtr<IFrameworkElement> actionsPanelAsFrameworkElement;
        RETURN_IF_FAILED(actionsPanel.As(&actionsPanelAsFrameworkElement));
        RETURN_IF_FAILED(
            XamlHelpers::SetStyleFromResourceDictionary(renderContext, L"Adaptive.Actions", actionsPanelAsFrameworkElement.Get()));

        ComPtr<IStackPanel> actionSet =
            XamlHelpers::CreateXamlClass<IStackPanel>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_StackPanel));
        ComPtr<IPanel> actionSetAsPanel;
        actionSet.As(&actionSetAsPanel);

        // Add buttons and show cards to panel
        XamlHelpers::AppendXamlElementToPanel(actionsPanel.Get(), actionSetAsPanel.Get());
        XamlHelpers::AppendXamlElementToPanel(showCardsStackPanel.Get(), actionSetAsPanel.Get());

        return actionSetAsPanel.CopyTo(actionSetControl);
    }

    HRESULT XamlBuilder::BuildActions(_In_ IAdaptiveCard* adaptiveCard,
                                      _In_ IVector<IAdaptiveActionElement*>* children,
                                      _In_ IPanel* bodyPanel,
                                      bool insertSeparator,
                                      _In_ IAdaptiveRenderContext* renderContext,
                                      _In_ ABI::AdaptiveNamespace::IAdaptiveRenderArgs* renderArgs)
    {
        ComPtr<IAdaptiveHostConfig> hostConfig;
        RETURN_IF_FAILED(renderContext->get_HostConfig(&hostConfig));
        ComPtr<IAdaptiveActionsConfig> actionsConfig;
        RETURN_IF_FAILED(hostConfig->get_Actions(actionsConfig.GetAddressOf()));

        // Create a separator between the body and the actions
        if (insertSeparator)
        {
            ABI::AdaptiveNamespace::Spacing spacing;
            RETURN_IF_FAILED(actionsConfig->get_Spacing(&spacing));

            UINT spacingSize;
            RETURN_IF_FAILED(GetSpacingSizeFromSpacing(hostConfig.Get(), spacing, &spacingSize));

            ABI::Windows::UI::Color color = {0};
            auto separator = XamlHelpers::CreateSeparator(renderContext, spacingSize, 0, color);
            XamlHelpers::AppendXamlElementToPanel(separator.Get(), bodyPanel);
        }

        ComPtr<IUIElement> actionSetControl;
        RETURN_IF_FAILED(BuildActionSetHelper(adaptiveCard, nullptr, children, renderContext, renderArgs, &actionSetControl));

        XamlHelpers::AppendXamlElementToPanel(actionSetControl.Get(), bodyPanel);
        return S_OK;
    }

    HRESULT AdaptiveActionSetRenderer::Render(_In_ IAdaptiveCardElement* adaptiveCardElement,
                                              _In_ IAdaptiveRenderContext* renderContext,
                                              _In_ IAdaptiveRenderArgs* renderArgs,
                                              _COM_Outptr_ IUIElement** actionSetControl) noexcept try
    {
        ComPtr<IAdaptiveHostConfig> hostConfig;
        RETURN_IF_FAILED(renderContext->get_HostConfig(&hostConfig));

        if (!XamlBuilder::SupportsInteractivity(hostConfig.Get()))
        {
            renderContext->AddWarning(
                ABI::AdaptiveNamespace::WarningStatusCode::InteractivityNotSupported,
                HStringReference(L"ActionSet was stripped from card because interactivity is not supported").Get());
            return S_OK;
        }

        ComPtr<IAdaptiveCardElement> cardElement(adaptiveCardElement);
        ComPtr<IAdaptiveActionSet> adaptiveActionSet;
        RETURN_IF_FAILED(cardElement.As(&adaptiveActionSet));

        ComPtr<IVector<IAdaptiveActionElement*>> actions;
        RETURN_IF_FAILED(adaptiveActionSet->get_Actions(&actions));

        return BuildActionSetHelper(nullptr, adaptiveActionSet.Get(), actions.Get(), renderContext, renderArgs, actionSetControl);
    }
    CATCH_RETURN;
}
