// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
#include "pch.h"

#include "AdaptiveActionSet.h"
#include "AdaptiveActionSetRenderer.h"
#include "AdaptiveElementParserRegistration.h"
#include "AdaptiveImage.h"
#include "AdaptiveRenderArgs.h"

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

    HRESULT AdaptiveActionSetRenderer::Render(_In_ IAdaptiveCardElement* cardElement,
                                              _In_ IAdaptiveRenderContext* renderContext,
                                              _In_ IAdaptiveRenderArgs* renderArgs,
                                              _COM_Outptr_ ABI::Windows::UI::Xaml::IUIElement** result) noexcept try
    {
        return XamlBuilder::BuildActionSet(cardElement, renderContext, renderArgs, result);
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
            auto separator = CreateSeparator(renderContext, spacingSize, 0, color);
            XamlHelpers::AppendXamlElementToPanel(separator.Get(), bodyPanel);
        }

        ComPtr<IUIElement> actionSetControl;
        RETURN_IF_FAILED(BuildActionSetHelper(adaptiveCard, nullptr, children, renderContext, renderArgs, &actionSetControl));

        XamlHelpers::AppendXamlElementToPanel(actionSetControl.Get(), bodyPanel);
        return S_OK;
    }

    static HRESULT GetButtonMargin(_In_ IAdaptiveActionsConfig* actionsConfig, Thickness &buttonMargin) noexcept
    {
        buttonMargin = {0, 0, 0, 0};
        UINT32 buttonSpacing;
        RETURN_IF_FAILED(actionsConfig->get_ButtonSpacing(&buttonSpacing));

        ABI::AdaptiveNamespace::ActionsOrientation actionsOrientation;
        RETURN_IF_FAILED(actionsConfig->get_ActionsOrientation(&actionsOrientation));

        if (actionsOrientation == ABI::AdaptiveNamespace::ActionsOrientation::Horizontal)
        {
            buttonMargin.Left = buttonMargin.Right = buttonSpacing / 2;
        }
        else
        {
            buttonMargin.Top = buttonMargin.Bottom = buttonSpacing / 2;
        }

        return S_OK;
    }

    HRESULT XamlBuilder::BuildAction(_In_ IAdaptiveActionElement* adaptiveActionElement,
                                     _In_ IAdaptiveRenderContext* renderContext,
                                     _In_ IAdaptiveRenderArgs* renderArgs,
                                     _Outptr_ IUIElement** actionControl)
    {
        // Render a button for the action
        ComPtr<IAdaptiveActionElement> action(adaptiveActionElement);
        ComPtr<IButton> button =
            XamlHelpers::CreateXamlClass<IButton>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_Button));

        ComPtr<IFrameworkElement> buttonFrameworkElement;
        RETURN_IF_FAILED(button.As(&buttonFrameworkElement));

        ComPtr<IAdaptiveHostConfig> hostConfig;
        RETURN_IF_FAILED(renderContext->get_HostConfig(&hostConfig));
        ComPtr<IAdaptiveActionsConfig> actionsConfig;
        RETURN_IF_FAILED(hostConfig->get_Actions(actionsConfig.GetAddressOf()));

        Thickness buttonMargin;
        RETURN_IF_FAILED(GetButtonMargin(actionsConfig.Get(), buttonMargin));
        RETURN_IF_FAILED(buttonFrameworkElement->put_Margin(buttonMargin));

        ABI::AdaptiveNamespace::ActionsOrientation actionsOrientation;
        RETURN_IF_FAILED(actionsConfig->get_ActionsOrientation(&actionsOrientation));

        ABI::AdaptiveNamespace::ActionAlignment actionAlignment;
        RETURN_IF_FAILED(actionsConfig->get_ActionAlignment(&actionAlignment));

        if (actionsOrientation == ABI::AdaptiveNamespace::ActionsOrientation::Horizontal)
        {
            // For horizontal alignment, we always use stretch
            RETURN_IF_FAILED(buttonFrameworkElement->put_HorizontalAlignment(
                ABI::Windows::UI::Xaml::HorizontalAlignment::HorizontalAlignment_Stretch));
        }
        else
        {
            switch (actionAlignment)
            {
            case ABI::AdaptiveNamespace::ActionAlignment::Center:
                RETURN_IF_FAILED(buttonFrameworkElement->put_HorizontalAlignment(ABI::Windows::UI::Xaml::HorizontalAlignment_Center));
                break;
            case ABI::AdaptiveNamespace::ActionAlignment::Left:
                RETURN_IF_FAILED(buttonFrameworkElement->put_HorizontalAlignment(ABI::Windows::UI::Xaml::HorizontalAlignment_Left));
                break;
            case ABI::AdaptiveNamespace::ActionAlignment::Right:
                RETURN_IF_FAILED(buttonFrameworkElement->put_HorizontalAlignment(ABI::Windows::UI::Xaml::HorizontalAlignment_Right));
                break;
            case ABI::AdaptiveNamespace::ActionAlignment::Stretch:
                RETURN_IF_FAILED(buttonFrameworkElement->put_HorizontalAlignment(ABI::Windows::UI::Xaml::HorizontalAlignment_Stretch));
                break;
            }
        }

        ABI::AdaptiveNamespace::ContainerStyle containerStyle;
        renderArgs->get_ContainerStyle(&containerStyle);

        boolean allowAboveTitleIconPlacement;
        RETURN_IF_FAILED(renderArgs->get_AllowAboveTitleIconPlacement(&allowAboveTitleIconPlacement));

        ArrangeButtonContent(action.Get(),
                             actionsConfig.Get(),
                             renderContext,
                             containerStyle,
                             hostConfig.Get(),
                             allowAboveTitleIconPlacement,
                             button.Get());

        ABI::AdaptiveNamespace::ActionType actionType;
        RETURN_IF_FAILED(action->get_ActionType(&actionType));

        ComPtr<IAdaptiveShowCardActionConfig> showCardActionConfig;
        RETURN_IF_FAILED(actionsConfig->get_ShowCard(&showCardActionConfig));
        ABI::AdaptiveNamespace::ActionMode showCardActionMode;
        RETURN_IF_FAILED(showCardActionConfig->get_ActionMode(&showCardActionMode));
        std::shared_ptr<std::vector<ComPtr<IUIElement>>> allShowCards = std::make_shared<std::vector<ComPtr<IUIElement>>>();

        // Add click handler which calls IAdaptiveActionInvoker::SendActionEvent
        ComPtr<IButtonBase> buttonBase;
        RETURN_IF_FAILED(button.As(&buttonBase));

        ComPtr<IAdaptiveActionInvoker> actionInvoker;
        RETURN_IF_FAILED(renderContext->get_ActionInvoker(&actionInvoker));
        EventRegistrationToken clickToken;
        RETURN_IF_FAILED(buttonBase->add_Click(Callback<IRoutedEventHandler>(
                                                   [action, actionInvoker](IInspectable* /*sender*/, IRoutedEventArgs * /*args*/) -> HRESULT {
                                                       return actionInvoker->SendActionEvent(action.Get());
                                                   })
                                                   .Get(),
                                               &clickToken));

        RETURN_IF_FAILED(HandleActionStyling(adaptiveActionElement, buttonFrameworkElement.Get(), renderContext));

        ComPtr<IUIElement> buttonAsUIElement;
        RETURN_IF_FAILED(button.As(&buttonAsUIElement));
        *actionControl = buttonAsUIElement.Detach();
        return S_OK;
    }

    HRESULT XamlBuilder::HandleActionStyling(_In_ IAdaptiveActionElement* adaptiveActionElement,
                                             _In_ IFrameworkElement* buttonFrameworkElement,
                                             _In_ IAdaptiveRenderContext* renderContext)
    {
        HString actionSentiment;
        RETURN_IF_FAILED(adaptiveActionElement->get_Style(actionSentiment.GetAddressOf()));

        INT32 isSentimentPositive{}, isSentimentDestructive{}, isSentimentDefault{};

        ComPtr<IResourceDictionary> resourceDictionary;
        RETURN_IF_FAILED(renderContext->get_OverrideStyles(&resourceDictionary));
        ComPtr<IStyle> styleToApply;

        ComPtr<AdaptiveNamespace::AdaptiveRenderContext> contextImpl =
            PeekInnards<AdaptiveNamespace::AdaptiveRenderContext>(renderContext);

        if ((SUCCEEDED(WindowsCompareStringOrdinal(actionSentiment.Get(), HStringReference(L"default").Get(), &isSentimentDefault)) &&
             (isSentimentDefault == 0)) ||
            WindowsIsStringEmpty(actionSentiment.Get()))
        {
            RETURN_IF_FAILED(SetStyleFromResourceDictionary(renderContext, L"Adaptive.Action", buttonFrameworkElement));
        }
        else if (SUCCEEDED(WindowsCompareStringOrdinal(actionSentiment.Get(), HStringReference(L"positive").Get(), &isSentimentPositive)) &&
                 (isSentimentPositive == 0))
        {
            if (SUCCEEDED(TryGetResourceFromResourceDictionaries<IStyle>(resourceDictionary.Get(), L"Adaptive.Action.Positive", &styleToApply)))
            {
                RETURN_IF_FAILED(buttonFrameworkElement->put_Style(styleToApply.Get()));
            }
            else
            {
                // By default, set the action background color to accent color
                ComPtr<IResourceDictionary> actionSentimentDictionary = contextImpl->GetDefaultActionSentimentDictionary();

                if (SUCCEEDED(TryGetResourceFromResourceDictionaries(actionSentimentDictionary.Get(),
                                                                     L"PositiveActionDefaultStyle",
                                                                     styleToApply.GetAddressOf())))
                {
                    RETURN_IF_FAILED(buttonFrameworkElement->put_Style(styleToApply.Get()));
                }
            }
        }
        else if (SUCCEEDED(WindowsCompareStringOrdinal(actionSentiment.Get(), HStringReference(L"destructive").Get(), &isSentimentDestructive)) &&
                 (isSentimentDestructive == 0))
        {
            if (SUCCEEDED(TryGetResourceFromResourceDictionaries<IStyle>(resourceDictionary.Get(), L"Adaptive.Action.Destructive", &styleToApply)))
            {
                RETURN_IF_FAILED(buttonFrameworkElement->put_Style(styleToApply.Get()));
            }
            else
            {
                // By default, set the action text color to attention color
                ComPtr<IResourceDictionary> actionSentimentDictionary = contextImpl->GetDefaultActionSentimentDictionary();

                if (SUCCEEDED(TryGetResourceFromResourceDictionaries(actionSentimentDictionary.Get(),
                                                                     L"DestructiveActionDefaultStyle",
                                                                     styleToApply.GetAddressOf())))
                {
                    RETURN_IF_FAILED(buttonFrameworkElement->put_Style(styleToApply.Get()));
                }
            }
        }
        else
        {
            HString actionSentimentStyle;
            RETURN_IF_FAILED(WindowsConcatString(HStringReference(L"Adaptive.Action.").Get(),
                                                 actionSentiment.Get(),
                                                 actionSentimentStyle.GetAddressOf()));
            RETURN_IF_FAILED(SetStyleFromResourceDictionary(renderContext,
                                                            StringToWstring(HStringToUTF8(actionSentimentStyle.Get())),
                                                            buttonFrameworkElement));
        }
        return S_OK;
    }

    HRESULT XamlBuilder::BuildActionSet(_In_ IAdaptiveCardElement* adaptiveCardElement,
                                        _In_ IAdaptiveRenderContext* renderContext,
                                        _In_ IAdaptiveRenderArgs* renderArgs,
                                        _COM_Outptr_ IUIElement** actionSetControl)
    {
        ComPtr<IAdaptiveHostConfig> hostConfig;
        RETURN_IF_FAILED(renderContext->get_HostConfig(&hostConfig));

        if (!SupportsInteractivity(hostConfig.Get()))
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

    HRESULT XamlBuilder::BuildActionSetHelper(_In_opt_ ABI::AdaptiveNamespace::IAdaptiveCard* adaptiveCard,
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
        RETURN_IF_FAILED(GetButtonMargin(actionsConfig.Get(), buttonMargin));
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
                            RETURN_IF_FAILED(WarnForFallbackDrop(renderContext, actionTypeString.Get()));
                            return S_OK;
                        }

                        case ABI::AdaptiveNamespace::FallbackType::Content:
                        {
                            ComPtr<IAdaptiveActionElement> actionFallback;
                            RETURN_IF_FAILED(action->get_FallbackContent(&actionFallback));

                            HString fallbackTypeString;
                            RETURN_IF_FAILED(actionFallback->get_ActionTypeString(fallbackTypeString.GetAddressOf()));
                            RETURN_IF_FAILED(
                                WarnForFallbackContentElement(renderContext, actionTypeString.Get(), fallbackTypeString.Get()));

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

                        RETURN_IF_FAILED(BuildShowCard(
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
            SetStyleFromResourceDictionary(renderContext, L"Adaptive.Actions", actionsPanelAsFrameworkElement.Get()));

        ComPtr<IStackPanel> actionSet =
            XamlHelpers::CreateXamlClass<IStackPanel>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_StackPanel));
        ComPtr<IPanel> actionSetAsPanel;
        actionSet.As(&actionSetAsPanel);

        // Add buttons and show cards to panel
        XamlHelpers::AppendXamlElementToPanel(actionsPanel.Get(), actionSetAsPanel.Get());
        XamlHelpers::AppendXamlElementToPanel(showCardsStackPanel.Get(), actionSetAsPanel.Get());

        return actionSetAsPanel.CopyTo(actionSetControl);
    }

    void XamlBuilder::ArrangeButtonContent(_In_ IAdaptiveActionElement* action,
                                           _In_ IAdaptiveActionsConfig* actionsConfig,
                                           _In_ IAdaptiveRenderContext* renderContext,
                                           ABI::AdaptiveNamespace::ContainerStyle containerStyle,
                                           _In_ ABI::AdaptiveNamespace::IAdaptiveHostConfig* hostConfig,
                                           bool allActionsHaveIcons,
                                           _In_ IButton* button)
    {
        HString title;
        THROW_IF_FAILED(action->get_Title(title.GetAddressOf()));

        HSTRING iconUrl;
        THROW_IF_FAILED(action->get_IconUrl(&iconUrl));

        ComPtr<IButton> localButton(button);

        // Check if the button has an iconUrl
        if (iconUrl != nullptr)
        {
            // Get icon configs
            ABI::AdaptiveNamespace::IconPlacement iconPlacement;
            UINT32 iconSize;

            THROW_IF_FAILED(actionsConfig->get_IconPlacement(&iconPlacement));
            THROW_IF_FAILED(actionsConfig->get_IconSize(&iconSize));

            // Define the alignment for the button contents
            ComPtr<IStackPanel> buttonContentsStackPanel =
                XamlHelpers::CreateXamlClass<IStackPanel>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_StackPanel));

            // Create image and add it to the button
            ComPtr<IAdaptiveImage> adaptiveImage;
            THROW_IF_FAILED(MakeAndInitialize<AdaptiveImage>(&adaptiveImage));

            adaptiveImage->put_Url(iconUrl);
            adaptiveImage->put_HorizontalAlignment(ABI::AdaptiveNamespace::HAlignment::Center);

            ComPtr<IAdaptiveCardElement> adaptiveCardElement;
            THROW_IF_FAILED(adaptiveImage.As(&adaptiveCardElement));
            ComPtr<AdaptiveRenderArgs> childRenderArgs;
            THROW_IF_FAILED(
                MakeAndInitialize<AdaptiveRenderArgs>(&childRenderArgs, containerStyle, buttonContentsStackPanel.Get(), nullptr));

            ComPtr<IAdaptiveElementRendererRegistration> elementRenderers;
            THROW_IF_FAILED(renderContext->get_ElementRenderers(&elementRenderers));

            ComPtr<IUIElement> buttonIcon;
            ComPtr<IAdaptiveElementRenderer> elementRenderer;
            THROW_IF_FAILED(elementRenderers->Get(HStringReference(L"Image").Get(), &elementRenderer));
            if (elementRenderer != nullptr)
            {
                elementRenderer->Render(adaptiveCardElement.Get(), renderContext, childRenderArgs.Get(), &buttonIcon);
                if (buttonIcon == nullptr)
                {
                    XamlHelpers::SetContent(localButton.Get(), title.Get());
                    return;
                }
            }

            // Create title text block
            ComPtr<ITextBlock> buttonText =
                XamlHelpers::CreateXamlClass<ITextBlock>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_TextBlock));
            THROW_IF_FAILED(buttonText->put_Text(title.Get()));
            THROW_IF_FAILED(buttonText->put_TextAlignment(TextAlignment::TextAlignment_Center));

            // Handle different arrangements inside button
            ComPtr<IFrameworkElement> buttonIconAsFrameworkElement;
            THROW_IF_FAILED(buttonIcon.As(&buttonIconAsFrameworkElement));
            ComPtr<IUIElement> separator;
            if (iconPlacement == ABI::AdaptiveNamespace::IconPlacement::AboveTitle && allActionsHaveIcons)
            {
                THROW_IF_FAILED(buttonContentsStackPanel->put_Orientation(Orientation::Orientation_Vertical));

                // Set icon height to iconSize (aspect ratio is automatically maintained)
                THROW_IF_FAILED(buttonIconAsFrameworkElement->put_Height(iconSize));
            }
            else
            {
                THROW_IF_FAILED(buttonContentsStackPanel->put_Orientation(Orientation::Orientation_Horizontal));

                // Add event to the image to resize itself when the textblock is rendered
                ComPtr<IImage> buttonIconAsImage;
                THROW_IF_FAILED(buttonIcon.As(&buttonIconAsImage));

                EventRegistrationToken eventToken;
                THROW_IF_FAILED(buttonIconAsImage->add_ImageOpened(
                    Callback<IRoutedEventHandler>([buttonIconAsFrameworkElement,
                                                   buttonText](IInspectable* /*sender*/, IRoutedEventArgs * /*args*/) -> HRESULT {
                        ComPtr<IFrameworkElement> buttonTextAsFrameworkElement;
                        RETURN_IF_FAILED(buttonText.As(&buttonTextAsFrameworkElement));

                        return SetMatchingHeight(buttonIconAsFrameworkElement.Get(), buttonTextAsFrameworkElement.Get());
                    })
                        .Get(),
                    &eventToken));

                // Only add spacing when the icon must be located at the left of the title
                UINT spacingSize;
                THROW_IF_FAILED(GetSpacingSizeFromSpacing(hostConfig, ABI::AdaptiveNamespace::Spacing::Default, &spacingSize));

                ABI::Windows::UI::Color color = {0};
                separator = CreateSeparator(renderContext, spacingSize, spacingSize, color, false);
            }

            ComPtr<IPanel> buttonContentsPanel;
            THROW_IF_FAILED(buttonContentsStackPanel.As(&buttonContentsPanel));

            // Add image to stack panel
            XamlHelpers::AppendXamlElementToPanel(buttonIcon.Get(), buttonContentsPanel.Get());

            // Add separator to stack panel
            if (separator != nullptr)
            {
                XamlHelpers::AppendXamlElementToPanel(separator.Get(), buttonContentsPanel.Get());
            }

            // Add text to stack panel
            XamlHelpers::AppendXamlElementToPanel(buttonText.Get(), buttonContentsPanel.Get());

            // Finally, put the stack panel inside the final button
            ComPtr<IContentControl> buttonContentControl;
            THROW_IF_FAILED(localButton.As(&buttonContentControl));
            THROW_IF_FAILED(buttonContentControl->put_Content(buttonContentsPanel.Get()));
        }
        else
        {
            XamlHelpers::SetContent(localButton.Get(), title.Get());
        }
    }

}
