// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
#include "pch.h"

#include "AdaptiveBase64Util.h"
#include "AdaptiveCardGetResourceStreamArgs.h"
#include "AdaptiveCardRendererComponent.h"
#include "AdaptiveCardResourceResolvers.h"
#include "AdaptiveColorsConfig.h"
#include "AdaptiveColorConfig.h"
#include "AdaptiveFeatureRegistration.h"
#include "AdaptiveHostConfig.h"
#include "AdaptiveImage.h"
#include "AdaptiveRenderArgs.h"
#include "AdaptiveShowCardAction.h"
#include "AdaptiveTextRun.h"
#include "DateTimeParser.h"
#include "ElementTagContent.h"
#include "FeatureRegistration.h"
#include "TextHelpers.h"
#include "json/json.h"
#include "MarkDownParser.h"
#include "MediaHelpers.h"
#include <robuffer.h>
#include "TileControl.h"
#include "WholeItemsPanel.h"
#include <windows.web.http.h>
#include <windows.web.http.filters.h>
#include "XamlBuilder.h"
#include "XamlHelpers.h"

using namespace Microsoft::WRL;
using namespace Microsoft::WRL::Wrappers;
using namespace ABI::AdaptiveNamespace;
using namespace ABI::Windows::Data::Json;
using namespace ABI::Windows::Foundation;
using namespace ABI::Windows::Foundation::Collections;
using namespace ABI::Windows::Storage;
using namespace ABI::Windows::Storage::Streams;
using namespace ABI::Windows::UI;
using namespace ABI::Windows::UI::Text;
using namespace ABI::Windows::UI::Xaml;
using namespace ABI::Windows::UI::Xaml::Data;
using namespace ABI::Windows::UI::Xaml::Documents;
using namespace ABI::Windows::UI::Xaml::Controls;
using namespace ABI::Windows::UI::Xaml::Controls::Primitives;
using namespace ABI::Windows::UI::Xaml::Markup;
using namespace ABI::Windows::UI::Xaml::Media;
using namespace ABI::Windows::UI::Xaml::Media::Imaging;
using namespace ABI::Windows::UI::Xaml::Shapes;
using namespace ABI::Windows::UI::Xaml::Input;
using namespace ABI::Windows::UI::Xaml::Automation;
using namespace ABI::Windows::Web::Http;
using namespace ABI::Windows::Web::Http::Filters;

constexpr PCWSTR c_BackgroundImageOverlayBrushKey = L"AdaptiveCard.BackgroundOverlayBrush";

namespace AdaptiveNamespace
{
    XamlBuilder::XamlBuilder()
    {
        m_imageLoadTracker.AddListener(dynamic_cast<IImageLoadTrackerListener*>(this));

        THROW_IF_FAILED(GetActivationFactory(HStringReference(RuntimeClass_Windows_Storage_Streams_RandomAccessStream).Get(),
                                             &m_randomAccessStreamStatics));
    }

    ComPtr<IUIElement> XamlBuilder::CreateSeparator(_In_ IAdaptiveRenderContext* renderContext,
                                                    UINT spacing,
                                                    UINT separatorThickness,
                                                    ABI::Windows::UI::Color separatorColor,
                                                    bool isHorizontal)
    {
        ComPtr<IGrid> separator =
            XamlHelpers::CreateXamlClass<IGrid>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_Grid));
        ComPtr<IFrameworkElement> separatorAsFrameworkElement;
        THROW_IF_FAILED(separator.As(&separatorAsFrameworkElement));

        ComPtr<IBrush> lineColorBrush = XamlHelpers::GetSolidColorBrush(separatorColor);
        ComPtr<IPanel> separatorAsPanel;
        THROW_IF_FAILED(separator.As(&separatorAsPanel));
        separatorAsPanel->put_Background(lineColorBrush.Get());

        UINT32 separatorMarginValue = spacing > separatorThickness ? (spacing - separatorThickness) / 2 : 0;
        Thickness margin = {0, 0, 0, 0};

        if (isHorizontal)
        {
            margin.Top = margin.Bottom = separatorMarginValue;
            separatorAsFrameworkElement->put_Height(separatorThickness);
        }
        else
        {
            margin.Left = margin.Right = separatorMarginValue;
            separatorAsFrameworkElement->put_Width(separatorThickness);
        }

        THROW_IF_FAILED(separatorAsFrameworkElement->put_Margin(margin));

        THROW_IF_FAILED(SetStyleFromResourceDictionary(renderContext, L"Adaptive.Separator", separatorAsFrameworkElement.Get()));

        ComPtr<IUIElement> result;
        THROW_IF_FAILED(separator.As(&result));
        return result;
    }

    HRESULT XamlBuilder::AllImagesLoaded()
    {
        FireAllImagesLoaded();
        return S_OK;
    }

    HRESULT XamlBuilder::ImagesLoadingHadError()
    {
        FireImagesLoadingHadError();
        return S_OK;
    }

    HRESULT XamlBuilder::BuildXamlTreeFromAdaptiveCard(_In_ IAdaptiveCard* adaptiveCard,
                                                       _Outptr_ IFrameworkElement** xamlTreeRoot,
                                                       _In_ IAdaptiveRenderContext* renderContext,
                                                       ComPtr<XamlBuilder> xamlBuilder,
                                                       ABI::AdaptiveNamespace::ContainerStyle defaultContainerStyle)
    {
        *xamlTreeRoot = nullptr;
        if (adaptiveCard != nullptr)
        {
            ComPtr<IAdaptiveHostConfig> hostConfig;
            RETURN_IF_FAILED(renderContext->get_HostConfig(&hostConfig));
            ComPtr<IAdaptiveCardConfig> adaptiveCardConfig;
            RETURN_IF_FAILED(hostConfig->get_AdaptiveCard(&adaptiveCardConfig));

            boolean allowCustomStyle;
            RETURN_IF_FAILED(adaptiveCardConfig->get_AllowCustomStyle(&allowCustomStyle));

            ABI::AdaptiveNamespace::ContainerStyle containerStyle = defaultContainerStyle;
            if (allowCustomStyle)
            {
                ABI::AdaptiveNamespace::ContainerStyle cardStyle;
                RETURN_IF_FAILED(adaptiveCard->get_Style(&cardStyle));

                if (cardStyle != ABI::AdaptiveNamespace::ContainerStyle::None)
                {
                    containerStyle = cardStyle;
                }
            }
            ComPtr<IAdaptiveRenderArgs> renderArgs;
            RETURN_IF_FAILED(MakeAndInitialize<AdaptiveRenderArgs>(&renderArgs, containerStyle, nullptr, nullptr));

            ComPtr<IPanel> bodyElementContainer;
            ComPtr<IUIElement> rootElement =
                CreateRootCardElement(adaptiveCard, renderContext, renderArgs.Get(), xamlBuilder, &bodyElementContainer);
            ComPtr<IFrameworkElement> rootAsFrameworkElement;
            RETURN_IF_FAILED(rootElement.As(&rootAsFrameworkElement));

            UINT32 cardMinHeight{};
            RETURN_IF_FAILED(adaptiveCard->get_MinHeight(&cardMinHeight));
            if (cardMinHeight > 0)
            {
                RETURN_IF_FAILED(rootAsFrameworkElement->put_MinHeight(cardMinHeight));
            }

            ComPtr<IAdaptiveActionElement> selectAction;
            RETURN_IF_FAILED(adaptiveCard->get_SelectAction(&selectAction));

            // Create a new IUIElement pointer to house the root element decorated with select action
            ComPtr<IUIElement> rootSelectActionElement;
            HandleSelectAction(nullptr,
                               selectAction.Get(),
                               renderContext,
                               rootElement.Get(),
                               SupportsInteractivity(hostConfig.Get()),
                               true,
                               &rootSelectActionElement);
            RETURN_IF_FAILED(rootSelectActionElement.As(&rootAsFrameworkElement));

            // Enumerate the child items of the card and build xaml for them
            ComPtr<IVector<IAdaptiveCardElement*>> body;
            RETURN_IF_FAILED(adaptiveCard->get_Body(&body));
            ComPtr<IAdaptiveRenderArgs> bodyRenderArgs;
            RETURN_IF_FAILED(
                MakeAndInitialize<AdaptiveRenderArgs>(&bodyRenderArgs, containerStyle, rootAsFrameworkElement.Get(), nullptr));
            RETURN_IF_FAILED(
                BuildPanelChildren(body.Get(), bodyElementContainer.Get(), renderContext, bodyRenderArgs.Get(), [](IUIElement*) {}));

            ABI::AdaptiveNamespace::VerticalContentAlignment verticalContentAlignment;
            RETURN_IF_FAILED(adaptiveCard->get_VerticalContentAlignment(&verticalContentAlignment));
            XamlBuilder::SetVerticalContentAlignmentToChildren(bodyElementContainer.Get(), verticalContentAlignment);

            ComPtr<IVector<IAdaptiveActionElement*>> actions;
            RETURN_IF_FAILED(adaptiveCard->get_Actions(&actions));
            UINT32 actionsSize;
            RETURN_IF_FAILED(actions->get_Size(&actionsSize));
            if (actionsSize > 0)
            {
                if (SupportsInteractivity(hostConfig.Get()))
                {
                    unsigned int bodyCount;
                    RETURN_IF_FAILED(body->get_Size(&bodyCount));
                    BuildActions(adaptiveCard,
                                 actions.Get(),
                                 bodyElementContainer.Get(),
                                 bodyCount > 0,
                                 renderContext,
                                 renderArgs.Get());
                }
                else
                {
                    renderContext->AddWarning(
                        ABI::AdaptiveNamespace::WarningStatusCode::InteractivityNotSupported,
                        HStringReference(L"Actions collection was present in card, but interactivity is not supported").Get());
                }
            }

            boolean isInShowCard;
            renderArgs->get_IsInShowCard(&isInShowCard);

            if (!isInShowCard)
            {
                RETURN_IF_FAILED(SetStyleFromResourceDictionary(renderContext, L"Adaptive.Card", rootAsFrameworkElement.Get()));
            }
            else
            {
                RETURN_IF_FAILED(
                    SetStyleFromResourceDictionary(renderContext, L"Adaptive.ShowCard.Card", rootAsFrameworkElement.Get()));
            }

            RETURN_IF_FAILED(rootAsFrameworkElement.CopyTo(xamlTreeRoot));

            if (!isInShowCard && (xamlBuilder != nullptr))
            {
                if (xamlBuilder->m_listeners.size() == 0)
                {
                    // If we're done and no one's listening for the images to load, make sure
                    // any outstanding image loads are no longer tracked.
                    xamlBuilder->m_imageLoadTracker.AbandonOutstandingImages();
                }
                else if (xamlBuilder->m_imageLoadTracker.GetTotalImagesTracked() == 0)
                {
                    // If there are no images to track, fire the all images loaded
                    // event to signal the xaml is ready
                    xamlBuilder->FireAllImagesLoaded();
                }
            }
        }
        return S_OK;
    }

    HRESULT XamlBuilder::AddListener(_In_ IXamlBuilderListener* listener) noexcept try
    {
        if (m_listeners.find(listener) == m_listeners.end())
        {
            m_listeners.emplace(listener);
        }
        else
        {
            return E_INVALIDARG;
        }
        return S_OK;
    }
    CATCH_RETURN;

    HRESULT XamlBuilder::RemoveListener(_In_ IXamlBuilderListener* listener) noexcept try
    {
        if (m_listeners.find(listener) != m_listeners.end())
        {
            m_listeners.erase(listener);
        }
        else
        {
            return E_INVALIDARG;
        }
        return S_OK;
    }
    CATCH_RETURN;

    void XamlBuilder::SetFixedDimensions(UINT width, UINT height) noexcept
    {
        m_fixedDimensions = true;
        m_fixedWidth = width;
        m_fixedHeight = height;
    }

    void XamlBuilder::SetEnableXamlImageHandling(bool enableXamlImageHandling) noexcept
    {
        m_enableXamlImageHandling = enableXamlImageHandling;
    }

    template<typename T>
    HRESULT XamlBuilder::TryGetResourceFromResourceDictionaries(_In_ IResourceDictionary* resourceDictionary,
                                                                std::wstring resourceName,
                                                                _COM_Outptr_ T** style) noexcept
    {
        if (resourceDictionary == nullptr)
        {
            return E_INVALIDARG;
        }

        *style = nullptr;

        // Get a resource key for the requested style that we can use for ResourceDictionary Lookups
        ComPtr<IPropertyValueStatics> propertyValueStatics;
        RETURN_IF_FAILED(GetActivationFactory(HStringReference(RuntimeClass_Windows_Foundation_PropertyValue).Get(),
                                              &propertyValueStatics));

        ComPtr<IInspectable> resourceKey;
        RETURN_IF_FAILED(propertyValueStatics->CreateString(HStringReference(resourceName.c_str()).Get(), resourceKey.GetAddressOf()));

        // Search for the named resource
        ComPtr<IResourceDictionary> strongDictionary = resourceDictionary;
        ComPtr<IMap<IInspectable*, IInspectable*>> resourceDictionaryMap;

        boolean hasKey{};
        RETURN_IF_FAILED(strongDictionary.As(&resourceDictionaryMap));
        RETURN_IF_FAILED(resourceDictionaryMap->HasKey(resourceKey.Get(), &hasKey));
        if (hasKey)
        {
            ComPtr<IInspectable> dictionaryValue;
            RETURN_IF_FAILED(resourceDictionaryMap->Lookup(resourceKey.Get(), dictionaryValue.GetAddressOf()));

            ComPtr<T> resourceToReturn;
            RETURN_IF_FAILED(dictionaryValue.As(&resourceToReturn));
            RETURN_IF_FAILED(resourceToReturn.CopyTo(style));
        }

        return E_FAIL;
    }

    HRESULT XamlBuilder::TryInsertResourceToResourceDictionaries(_In_ IResourceDictionary* resourceDictionary,
                                                                 std::wstring resourceName,
                                                                 _In_ IInspectable* value)
    {
        if (resourceDictionary == nullptr)
        {
            return E_INVALIDARG;
        }

        try
        {
            ComPtr<IPropertyValueStatics> propertyValueStatics;
            THROW_IF_FAILED(GetActivationFactory(HStringReference(RuntimeClass_Windows_Foundation_PropertyValue).Get(),
                                                 &propertyValueStatics));

            ComPtr<IInspectable> resourceKey;
            THROW_IF_FAILED(propertyValueStatics->CreateString(HStringReference(resourceName.c_str()).Get(),
                                                               resourceKey.GetAddressOf()));

            ComPtr<IResourceDictionary> strongDictionary = resourceDictionary;
            ComPtr<IMap<IInspectable*, IInspectable*>> resourceDictionaryMap;
            THROW_IF_FAILED(strongDictionary.As(&resourceDictionaryMap));

            boolean replaced{};
            THROW_IF_FAILED(resourceDictionaryMap->Insert(resourceKey.Get(), value, &replaced));
            return S_OK;
        }
        catch (...)
        {
        }
        return E_FAIL;
    }

    HRESULT XamlBuilder::SetStyleFromResourceDictionary(_In_ IAdaptiveRenderContext* renderContext,
                                                        std::wstring resourceName,
                                                        _In_ IFrameworkElement* frameworkElement) noexcept
    {
        ComPtr<IResourceDictionary> resourceDictionary;
        RETURN_IF_FAILED(renderContext->get_OverrideStyles(&resourceDictionary));

        ComPtr<IStyle> style;
        if (SUCCEEDED(TryGetResourceFromResourceDictionaries<IStyle>(resourceDictionary.Get(), resourceName, &style)))
        {
            RETURN_IF_FAILED(frameworkElement->put_Style(style.Get()));
        }

        return S_OK;
    }

    static void ApplyMarginToXamlElement(_In_ IAdaptiveHostConfig* hostConfig, _In_ IFrameworkElement* element)
    {
        ComPtr<IFrameworkElement> localElement(element);
        ComPtr<IAdaptiveSpacingConfig> spacingConfig;
        THROW_IF_FAILED(hostConfig->get_Spacing(&spacingConfig));

        UINT32 padding;
        spacingConfig->get_Padding(&padding);
        Thickness margin = {(double)padding, (double)padding, (double)padding, (double)padding};

        THROW_IF_FAILED(localElement->put_Margin(margin));
    }

    ComPtr<IUIElement> XamlBuilder::CreateRootCardElement(_In_ IAdaptiveCard* adaptiveCard,
                                                          _In_ IAdaptiveRenderContext* renderContext,
                                                          _In_ IAdaptiveRenderArgs* renderArgs,
                                                          ComPtr<XamlBuilder> xamlBuilder,
                                                          _Outptr_ IPanel** bodyElementContainer)
    {
        // The root of an adaptive card is a composite of several elements, depending on the card
        // properties.  From back to front these are:
        // Grid - Root element, used to enable children to stack above each other and size to fit
        // Image (optional) - Holds the background image if one is set
        // Shape (optional) - Provides the background image overlay, if one is set
        // StackPanel - The container for all the card's body elements
        ComPtr<IGrid> rootElement =
            XamlHelpers::CreateXamlClass<IGrid>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_Grid));
        ComPtr<IAdaptiveHostConfig> hostConfig;
        THROW_IF_FAILED(renderContext->get_HostConfig(&hostConfig));
        ComPtr<IAdaptiveCardConfig> adaptiveCardConfig;
        THROW_IF_FAILED(hostConfig->get_AdaptiveCard(&adaptiveCardConfig));

        ComPtr<IPanel> rootAsPanel;
        THROW_IF_FAILED(rootElement.As(&rootAsPanel));
        ABI::AdaptiveNamespace::ContainerStyle containerStyle;
        THROW_IF_FAILED(renderArgs->get_ContainerStyle(&containerStyle));

        ABI::Windows::UI::Color backgroundColor;
        if (SUCCEEDED(GetBackgroundColorFromStyle(containerStyle, hostConfig.Get(), &backgroundColor)))
        {
            ComPtr<IBrush> backgroundColorBrush = XamlHelpers::GetSolidColorBrush(backgroundColor);
            THROW_IF_FAILED(rootAsPanel->put_Background(backgroundColorBrush.Get()));
        }

        ComPtr<IAdaptiveBackgroundImage> backgroundImage;
        BOOL backgroundImageIsValid;
        THROW_IF_FAILED(adaptiveCard->get_BackgroundImage(&backgroundImage));
        THROW_IF_FAILED(IsBackgroundImageValid(backgroundImage.Get(), &backgroundImageIsValid));
        if (backgroundImageIsValid)
        {
            ApplyBackgroundToRoot(rootAsPanel.Get(), backgroundImage.Get(), renderContext, renderArgs);
        }

        ComPtr<IAdaptiveSpacingConfig> spacingConfig;
        THROW_IF_FAILED(hostConfig->get_Spacing(&spacingConfig));

        UINT32 padding;
        THROW_IF_FAILED(spacingConfig->get_Padding(&padding));

        // Configure WholeItemsPanel to not clip bleeding containers
        WholeItemsPanel::SetBleedMargin(padding);

        // Now create the inner stack panel to serve as the root host for all the
        // body elements and apply padding from host configuration
        ComPtr<WholeItemsPanel> bodyElementHost;
        THROW_IF_FAILED(MakeAndInitialize<WholeItemsPanel>(&bodyElementHost));
        bodyElementHost->SetMainPanel(TRUE);
        bodyElementHost->SetAdaptiveHeight(TRUE);

        ComPtr<IFrameworkElement> bodyElementHostAsElement;
        THROW_IF_FAILED(bodyElementHost.As(&bodyElementHostAsElement));
        ApplyMarginToXamlElement(hostConfig.Get(), bodyElementHostAsElement.Get());

        ABI::AdaptiveNamespace::HeightType adaptiveCardHeightType;
        THROW_IF_FAILED(adaptiveCard->get_Height(&adaptiveCardHeightType));

        XamlHelpers::AppendXamlElementToPanel(bodyElementHost.Get(), rootAsPanel.Get(), adaptiveCardHeightType);
        THROW_IF_FAILED(bodyElementHost.CopyTo(bodyElementContainer));

        if (xamlBuilder && xamlBuilder->m_fixedDimensions)
        {
            ComPtr<IFrameworkElement> rootAsFrameworkElement;
            THROW_IF_FAILED(rootElement.As(&rootAsFrameworkElement));
            rootAsFrameworkElement->put_Width(xamlBuilder->m_fixedWidth);
            rootAsFrameworkElement->put_Height(xamlBuilder->m_fixedHeight);
            rootAsFrameworkElement->put_MaxHeight(xamlBuilder->m_fixedHeight);
        }

        if (adaptiveCardHeightType == ABI::AdaptiveNamespace::HeightType::Stretch)
        {
            ComPtr<IFrameworkElement> rootAsFrameworkElement;
            THROW_IF_FAILED(rootElement.As(&rootAsFrameworkElement));
            rootAsFrameworkElement->put_VerticalAlignment(ABI::Windows::UI::Xaml::VerticalAlignment::VerticalAlignment_Stretch);
        }

        ComPtr<IUIElement> rootAsUIElement;
        THROW_IF_FAILED(rootElement.As(&rootAsUIElement));
        return rootAsUIElement;
    }

    void XamlBuilder::ApplyBackgroundToRoot(_In_ IPanel* rootPanel,
                                            _In_ IAdaptiveBackgroundImage* backgroundImage,
                                            _In_ IAdaptiveRenderContext* renderContext,
                                            _In_ IAdaptiveRenderArgs* renderArgs)
    {
        // In order to reuse the image creation code paths, we simply create an adaptive card
        // image element and then build that into xaml and apply to the root.
        ComPtr<IAdaptiveImage> adaptiveImage;
        HSTRING url;
        THROW_IF_FAILED(MakeAndInitialize<AdaptiveImage>(&adaptiveImage));
        THROW_IF_FAILED(backgroundImage->get_Url(&url));
        THROW_IF_FAILED(adaptiveImage->put_Url(url));

        ComPtr<IAdaptiveCardElement> adaptiveCardElement;
        THROW_IF_FAILED(adaptiveImage.As(&adaptiveCardElement));

        ComPtr<IAdaptiveElementRendererRegistration> elementRenderers;
        THROW_IF_FAILED(renderContext->get_ElementRenderers(&elementRenderers));

        ComPtr<IAdaptiveElementRenderer> elementRenderer;
        THROW_IF_FAILED(elementRenderers->Get(HStringReference(L"Image").Get(), &elementRenderer));

        ComPtr<IUIElement> background;
        if (elementRenderer != nullptr)
        {
            elementRenderer->Render(adaptiveCardElement.Get(), renderContext, renderArgs, &background);
            if (background == nullptr)
            {
                return;
            }
        }

        ComPtr<IImage> xamlImage;
        THROW_IF_FAILED(background.As(&xamlImage));

        ABI::AdaptiveNamespace::BackgroundImageFillMode fillMode;
        THROW_IF_FAILED(backgroundImage->get_FillMode(&fillMode));

        // Creates the background image for all fill modes
        ComPtr<TileControl> tileControl;
        THROW_IF_FAILED(MakeAndInitialize<TileControl>(&tileControl));
        THROW_IF_FAILED(tileControl->put_BackgroundImage(backgroundImage));

        ComPtr<IFrameworkElement> rootElement;
        THROW_IF_FAILED(rootPanel->QueryInterface(rootElement.GetAddressOf()));
        THROW_IF_FAILED(tileControl->put_RootElement(rootElement.Get()));

        THROW_IF_FAILED(tileControl->LoadImageBrush(background.Get()));

        ComPtr<IFrameworkElement> backgroundAsFrameworkElement;
        THROW_IF_FAILED(tileControl.As(&backgroundAsFrameworkElement));

        XamlHelpers::AppendXamlElementToPanel(backgroundAsFrameworkElement.Get(), rootPanel);

        // The overlay applied to the background image is determined by a resouce, so create
        // the overlay if that resources exists
        ComPtr<IResourceDictionary> resourceDictionary;
        THROW_IF_FAILED(renderContext->get_OverrideStyles(&resourceDictionary));
        ComPtr<IBrush> backgroundOverlayBrush;
        if (SUCCEEDED(TryGetResourceFromResourceDictionaries<IBrush>(resourceDictionary.Get(), c_BackgroundImageOverlayBrushKey, &backgroundOverlayBrush)))
        {
            ComPtr<IShape> overlayRectangle =
                XamlHelpers::CreateXamlClass<IShape>(HStringReference(RuntimeClass_Windows_UI_Xaml_Shapes_Rectangle));
            THROW_IF_FAILED(overlayRectangle->put_Fill(backgroundOverlayBrush.Get()));

            ComPtr<IUIElement> overlayRectangleAsUIElement;
            THROW_IF_FAILED(overlayRectangle.As(&overlayRectangleAsUIElement));
            XamlHelpers::AppendXamlElementToPanel(overlayRectangle.Get(), rootPanel);
        }
    }

    void XamlBuilder::FireAllImagesLoaded()
    {
        for (auto& listener : m_listeners)
        {
            listener->AllImagesLoaded();
        }
    }

    void XamlBuilder::FireImagesLoadingHadError()
    {
        for (auto& listener : m_listeners)
        {
            listener->ImagesLoadingHadError();
        }
    }

    HRESULT XamlBuilder::AddRenderedControl(_In_ IUIElement* newControl,
                                            _In_ IAdaptiveCardElement* element,
                                            _In_ IPanel* parentPanel,
                                            _In_ IUIElement* separator,
                                            _In_ IColumnDefinition* columnDefinition,
                                            std::function<void(IUIElement* child)> childCreatedCallback)
    {
        if (newControl != nullptr)
        {
            boolean isVisible;
            RETURN_IF_FAILED(element->get_IsVisible(&isVisible));

            if (!isVisible)
            {
                RETURN_IF_FAILED(newControl->put_Visibility(Visibility_Collapsed));
            }

            ComPtr<IUIElement> localControl(newControl);
            ComPtr<IFrameworkElement> newControlAsFrameworkElement;
            RETURN_IF_FAILED(localControl.As(&newControlAsFrameworkElement));

            HString id;
            RETURN_IF_FAILED(element->get_Id(id.GetAddressOf()));

            if (id.IsValid())
            {
                RETURN_IF_FAILED(newControlAsFrameworkElement->put_Name(id.Get()));
            }

            ComPtr<ElementTagContent> tagContent;
            RETURN_IF_FAILED(MakeAndInitialize<ElementTagContent>(&tagContent, element, parentPanel, separator, columnDefinition, isVisible));
            RETURN_IF_FAILED(newControlAsFrameworkElement->put_Tag(tagContent.Get()));

            ABI::AdaptiveNamespace::HeightType heightType{};
            RETURN_IF_FAILED(element->get_Height(&heightType));
            XamlHelpers::AppendXamlElementToPanel(newControl, parentPanel, heightType);

            childCreatedCallback(newControl);
        }
        return S_OK;
    }

    void XamlBuilder::AddSeparatorIfNeeded(int& currentElement,
                                           _In_ IAdaptiveCardElement* element,
                                           _In_ IAdaptiveHostConfig* hostConfig,
                                           _In_ IAdaptiveRenderContext* renderContext,
                                           _In_ IPanel* parentPanel,
                                           _Outptr_ IUIElement** addedSeparator)
    {
        // First element does not need a separator added
        if (currentElement++ > 0)
        {
            bool needsSeparator;
            UINT spacing;
            UINT separatorThickness;
            ABI::Windows::UI::Color separatorColor;
            GetSeparationConfigForElement(element, hostConfig, &spacing, &separatorThickness, &separatorColor, &needsSeparator);
            if (needsSeparator)
            {
                auto separator = CreateSeparator(renderContext, spacing, separatorThickness, separatorColor);
                XamlHelpers::AppendXamlElementToPanel(separator.Get(), parentPanel);
                THROW_IF_FAILED(separator.CopyTo(addedSeparator));
            }
        }
    }

    HRESULT XamlBuilder::SetSeparatorVisibility(_In_ IPanel* parentPanel)
    {
        // Iterate over the elements in a container and ensure that the correct separators are marked as visible
        ComPtr<IVector<UIElement*>> children;
        RETURN_IF_FAILED(parentPanel->get_Children(&children));

        bool foundPreviousVisibleElement = false;
        XamlHelpers::IterateOverVector<UIElement, IUIElement>(children.Get(), [&](IUIElement* child) {
            ComPtr<IUIElement> localChild(child);

            ComPtr<IFrameworkElement> childAsFrameworkElement;
            RETURN_IF_FAILED(localChild.As(&childAsFrameworkElement));

            // Get the tag for the element. The separators themselves will not have tags.
            ComPtr<IInspectable> tag;
            RETURN_IF_FAILED(childAsFrameworkElement->get_Tag(&tag));

            if (tag)
            {
                ComPtr<IElementTagContent> elementTagContent;
                RETURN_IF_FAILED(tag.As(&elementTagContent));

                ComPtr<IUIElement> separator;
                RETURN_IF_FAILED(elementTagContent->get_Separator(&separator));

                Visibility visibility;
                RETURN_IF_FAILED(child->get_Visibility(&visibility));

                boolean expectedVisibility{};
                RETURN_IF_FAILED(elementTagContent->get_ExpectedVisibility(&expectedVisibility));

                if (separator)
                {
                    if (!expectedVisibility || !foundPreviousVisibleElement)
                    {
                        // If the element is collapsed, or if it's the first visible element, collapse the separator
                        // Images are hidden while they are retrieved, we shouldn't hide the separator
                        RETURN_IF_FAILED(separator->put_Visibility(Visibility_Collapsed));
                    }
                    else
                    {
                        // Otherwise show the separator
                        RETURN_IF_FAILED(separator->put_Visibility(Visibility_Visible));
                    }
                }

                foundPreviousVisibleElement |= (visibility == Visibility_Visible);
            }

            return S_OK;
        });

        return S_OK;
    }

    HRESULT XamlBuilder::RenderFallback(_In_ IAdaptiveCardElement* currentElement,
                                        _In_ IAdaptiveRenderContext* renderContext,
                                        _In_ IAdaptiveRenderArgs* renderArgs,
                                        _COM_Outptr_ IUIElement** result)
    {
        ComPtr<IAdaptiveElementRendererRegistration> elementRenderers;
        RETURN_IF_FAILED(renderContext->get_ElementRenderers(&elementRenderers));

        ABI::AdaptiveNamespace::FallbackType elementFallback;
        RETURN_IF_FAILED(currentElement->get_FallbackType(&elementFallback));

        HString elementType;
        RETURN_IF_FAILED(currentElement->get_ElementTypeString(elementType.GetAddressOf()));

        bool fallbackHandled = false;
        ComPtr<IUIElement> fallbackControl;
        switch (elementFallback)
        {
        case ABI::AdaptiveNamespace::FallbackType::Content:
        {
            // We have content, get the type of the fallback element
            ComPtr<IAdaptiveCardElement> fallbackElement;
            RETURN_IF_FAILED(currentElement->get_FallbackContent(&fallbackElement));

            HString fallbackElementType;
            RETURN_IF_FAILED(fallbackElement->get_ElementTypeString(fallbackElementType.GetAddressOf()));

            RETURN_IF_FAILED(WarnForFallbackContentElement(renderContext, elementType.Get(), fallbackElementType.Get()));

            // Try to render the fallback element
            ComPtr<IAdaptiveElementRenderer> fallbackElementRenderer;
            RETURN_IF_FAILED(elementRenderers->Get(fallbackElementType.Get(), &fallbackElementRenderer));
            HRESULT hr = E_PERFORM_FALLBACK;

            if (fallbackElementRenderer)
            {
                // perform this element's fallback
                hr = fallbackElementRenderer->Render(fallbackElement.Get(), renderContext, renderArgs, &fallbackControl);
            }

            if (hr == E_PERFORM_FALLBACK)
            {
                // The fallback content told us to fallback, make a recursive call to this method
                RETURN_IF_FAILED(RenderFallback(fallbackElement.Get(), renderContext, renderArgs, &fallbackControl));
            }
            else
            {
                // Check the non-fallback return value from the render call
                RETURN_IF_FAILED(hr);
            }

            // We handled the fallback content
            fallbackHandled = true;
            break;
        }
        case ABI::AdaptiveNamespace::FallbackType::Drop:
        {
            // If the fallback is drop, nothing to do but warn
            RETURN_IF_FAILED(WarnForFallbackDrop(renderContext, elementType.Get()));
            fallbackHandled = true;
            break;
        }
        case ABI::AdaptiveNamespace::FallbackType::None:
        default:
        {
            break;
        }
        }

        if (fallbackHandled)
        {
            // We did it, copy out the result if any
            RETURN_IF_FAILED(fallbackControl.CopyTo(result));
            return S_OK;
        }
        else
        {
            // We didn't do it, can our ancestor?
            boolean ancestorHasFallback;
            RETURN_IF_FAILED(renderArgs->get_AncestorHasFallback(&ancestorHasFallback));

            if (!ancestorHasFallback)
            {
                // standard unknown element handling
                std::wstring errorString = L"No Renderer found for type: ";
                errorString += elementType.GetRawBuffer(nullptr);
                RETURN_IF_FAILED(renderContext->AddWarning(ABI::AdaptiveNamespace::WarningStatusCode::NoRendererForType,
                                                           HStringReference(errorString.c_str()).Get()));
                return S_OK;
            }
            else
            {
                return E_PERFORM_FALLBACK;
            }
        }
    }

    HRESULT XamlBuilder::BuildPanelChildren(_In_ IVector<IAdaptiveCardElement*>* children,
                                            _In_ IPanel* parentPanel,
                                            _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                            _In_ ABI::AdaptiveNamespace::IAdaptiveRenderArgs* renderArgs,
                                            std::function<void(IUIElement* child)> childCreatedCallback) noexcept
    {
        int iElement = 0;
        unsigned int childrenSize;
        RETURN_IF_FAILED(children->get_Size(&childrenSize));
        boolean ancestorHasFallback;
        RETURN_IF_FAILED(renderArgs->get_AncestorHasFallback(&ancestorHasFallback));

        ComPtr<IAdaptiveFeatureRegistration> featureRegistration;
        RETURN_IF_FAILED(renderContext->get_FeatureRegistration(&featureRegistration));
        ComPtr<AdaptiveFeatureRegistration> featureRegistrationImpl = PeekInnards<AdaptiveFeatureRegistration>(featureRegistration);
        std::shared_ptr<FeatureRegistration> sharedFeatureRegistration = featureRegistrationImpl->GetSharedFeatureRegistration();

        HRESULT hr = XamlHelpers::IterateOverVectorWithFailure<IAdaptiveCardElement>(children, ancestorHasFallback, [&](IAdaptiveCardElement* element) {
            HRESULT hr = S_OK;

            // Get fallback state
            ABI::AdaptiveNamespace::FallbackType elementFallback;
            RETURN_IF_FAILED(element->get_FallbackType(&elementFallback));
            const bool elementHasFallback = (elementFallback != FallbackType_None);
            RETURN_IF_FAILED(renderArgs->put_AncestorHasFallback(elementHasFallback || ancestorHasFallback));

            // Check to see if element's requirements are being met
            boolean requirementsMet;
            RETURN_IF_FAILED(element->MeetsRequirements(featureRegistration.Get(), &requirementsMet));
            hr = requirementsMet ? S_OK : E_PERFORM_FALLBACK;

            // Get element renderer
            ComPtr<IAdaptiveElementRendererRegistration> elementRenderers;
            RETURN_IF_FAILED(renderContext->get_ElementRenderers(&elementRenderers));
            ComPtr<IAdaptiveElementRenderer> elementRenderer;
            HString elementType;
            RETURN_IF_FAILED(element->get_ElementTypeString(elementType.GetAddressOf()));
            RETURN_IF_FAILED(elementRenderers->Get(elementType.Get(), &elementRenderer));

            ComPtr<IAdaptiveHostConfig> hostConfig;
            RETURN_IF_FAILED(renderContext->get_HostConfig(&hostConfig));

            // If we have a renderer, render the element
            ComPtr<IUIElement> newControl;
            if (SUCCEEDED(hr) && elementRenderer != nullptr)
            {
                hr = elementRenderer->Render(element, renderContext, renderArgs, newControl.GetAddressOf());
            }

            // If we don't have a renderer, or if the renderer told us to perform fallback, try falling back
            if (elementRenderer == nullptr || hr == E_PERFORM_FALLBACK)
            {
                RETURN_IF_FAILED(RenderFallback(element, renderContext, renderArgs, &newControl));
            }

            // If we got a control, add a separator if needed and the control to the parent panel
            if (newControl != nullptr)
            {
                ComPtr<IUIElement> separator;
                AddSeparatorIfNeeded(iElement, element, hostConfig.Get(), renderContext, parentPanel, &separator);

                RETURN_IF_FAILED(AddRenderedControl(newControl.Get(), element, parentPanel, separator.Get(), nullptr, childCreatedCallback));
            }

            // Revert the ancestorHasFallback value
            renderArgs->put_AncestorHasFallback(ancestorHasFallback);
            return hr;
        });

        RETURN_IF_FAILED(SetSeparatorVisibility(parentPanel));
        return hr;
    }

    HRESULT XamlBuilder::HandleToggleVisibilityClick(_In_ IFrameworkElement* cardFrameworkElement, _In_ IAdaptiveActionElement* action)
    {
        ComPtr<IAdaptiveActionElement> localAction(action);
        ComPtr<IAdaptiveToggleVisibilityAction> toggleAction;
        RETURN_IF_FAILED(localAction.As(&toggleAction));

        ComPtr<IVector<AdaptiveToggleVisibilityTarget*>> targets;
        RETURN_IF_FAILED(toggleAction->get_TargetElements(&targets));

        ComPtr<IIterable<AdaptiveToggleVisibilityTarget*>> targetsIterable;
        RETURN_IF_FAILED(targets.As<IIterable<AdaptiveToggleVisibilityTarget*>>(&targetsIterable));

        boolean hasCurrent;
        ComPtr<IIterator<AdaptiveToggleVisibilityTarget*>> targetIterator;
        HRESULT hr = targetsIterable->First(&targetIterator);
        RETURN_IF_FAILED(targetIterator->get_HasCurrent(&hasCurrent));

        std::unordered_set<IPanel*> parentPanels;
        while (SUCCEEDED(hr) && hasCurrent)
        {
            ComPtr<IAdaptiveToggleVisibilityTarget> currentTarget;
            RETURN_IF_FAILED(targetIterator->get_Current(&currentTarget));

            HString toggleId;
            RETURN_IF_FAILED(currentTarget->get_ElementId(toggleId.GetAddressOf()));

            ABI::AdaptiveNamespace::IsVisible toggle;
            RETURN_IF_FAILED(currentTarget->get_IsVisible(&toggle));

            ComPtr<IInspectable> toggleElement;
            RETURN_IF_FAILED(cardFrameworkElement->FindName(toggleId.Get(), &toggleElement));

            if (toggleElement != nullptr)
            {
                ComPtr<IUIElement> toggleElementAsUIElement;
                RETURN_IF_FAILED(toggleElement.As(&toggleElementAsUIElement));

                ComPtr<IFrameworkElement> toggleElementAsFrameworkElement;
                RETURN_IF_FAILED(toggleElement.As(&toggleElementAsFrameworkElement));

                ComPtr<IInspectable> tag;
                RETURN_IF_FAILED(toggleElementAsFrameworkElement->get_Tag(&tag));

                ComPtr<IElementTagContent> elementTagContent;
                RETURN_IF_FAILED(tag.As(&elementTagContent));

                Visibility visibilityToSet = Visibility_Visible;
                if (toggle == ABI::AdaptiveNamespace::IsVisible_IsVisibleTrue)
                {
                    visibilityToSet = Visibility_Visible;
                }
                else if (toggle == ABI::AdaptiveNamespace::IsVisible_IsVisibleFalse)
                {
                    visibilityToSet = Visibility_Collapsed;
                }
                else if (toggle == ABI::AdaptiveNamespace::IsVisible_IsVisibleToggle)
                {
                    boolean currentVisibility{};
                    RETURN_IF_FAILED(elementTagContent->get_ExpectedVisibility(&currentVisibility));
                    visibilityToSet = (currentVisibility) ? Visibility_Collapsed : Visibility_Visible;
                }

                RETURN_IF_FAILED(toggleElementAsUIElement->put_Visibility(visibilityToSet));
                RETURN_IF_FAILED(elementTagContent->set_ExpectedVisibility(visibilityToSet == Visibility_Visible));

                ComPtr<IPanel> parentPanel;
                RETURN_IF_FAILED(elementTagContent->get_ParentPanel(&parentPanel));
                parentPanels.insert(parentPanel.Get());

                ComPtr<IAdaptiveCardElement> cardElement;
                RETURN_IF_FAILED(elementTagContent->get_AdaptiveCardElement(&cardElement));

                // If the element we're toggling is a column, we'll need to change the width on the column definition
                ComPtr<IAdaptiveColumn> cardElementAsColumn;
                if (SUCCEEDED(cardElement.As(&cardElementAsColumn)))
                {
                    ComPtr<IColumnDefinition> columnDefinition;
                    RETURN_IF_FAILED(elementTagContent->get_ColumnDefinition(&columnDefinition));
                    RETURN_IF_FAILED(HandleColumnWidth(cardElementAsColumn.Get(),
                                                       (visibilityToSet == Visibility_Visible),
                                                       columnDefinition.Get()));
                }
            }

            hr = targetIterator->MoveNext(&hasCurrent);
        }

        for (auto parentPanel : parentPanels)
        {
            SetSeparatorVisibility(parentPanel);
        }

        return S_OK;
    }

    void XamlBuilder::GetSeparationConfigForElement(_In_ IAdaptiveCardElement* cardElement,
                                                    _In_ IAdaptiveHostConfig* hostConfig,
                                                    _Out_ UINT* spacing,
                                                    _Out_ UINT* separatorThickness,
                                                    _Out_ ABI::Windows::UI::Color* separatorColor,
                                                    _Out_ bool* needsSeparator)
    {
        ABI::AdaptiveNamespace::Spacing elementSpacing;
        THROW_IF_FAILED(cardElement->get_Spacing(&elementSpacing));

        UINT localSpacing;
        THROW_IF_FAILED(GetSpacingSizeFromSpacing(hostConfig, elementSpacing, &localSpacing));

        boolean hasSeparator;
        THROW_IF_FAILED(cardElement->get_Separator(&hasSeparator));

        ABI::Windows::UI::Color localColor = {0};
        UINT localThickness = 0;
        if (hasSeparator)
        {
            ComPtr<IAdaptiveSeparatorConfig> separatorConfig;
            THROW_IF_FAILED(hostConfig->get_Separator(&separatorConfig));

            THROW_IF_FAILED(separatorConfig->get_LineColor(&localColor));
            THROW_IF_FAILED(separatorConfig->get_LineThickness(&localThickness));
        }

        *needsSeparator = hasSeparator || (elementSpacing != ABI::AdaptiveNamespace::Spacing::None);

        *spacing = localSpacing;
        *separatorThickness = localThickness;
        *separatorColor = localColor;
    }

    HRESULT XamlBuilder::SetAutoImageSize(_In_ IFrameworkElement* imageControl,
                                          _In_ IInspectable* parentElement,
                                          _In_ IBitmapSource* imageSource,
                                          bool setVisible)
    {
        INT32 pixelHeight;
        RETURN_IF_FAILED(imageSource->get_PixelHeight(&pixelHeight));
        INT32 pixelWidth;
        RETURN_IF_FAILED(imageSource->get_PixelWidth(&pixelWidth));
        DOUBLE maxHeight;
        DOUBLE maxWidth;
        ComPtr<IInspectable> localParentElement(parentElement);
        ComPtr<IFrameworkElement> localElement(imageControl);
        ComPtr<IColumnDefinition> parentAsColumnDefinition;

        RETURN_IF_FAILED(localElement->get_MaxHeight(&maxHeight));
        RETURN_IF_FAILED(localElement->get_MaxWidth(&maxWidth));

        if (SUCCEEDED(localParentElement.As(&parentAsColumnDefinition)))
        {
            DOUBLE parentWidth;
            RETURN_IF_FAILED(parentAsColumnDefinition->get_ActualWidth(&parentWidth));
            if (parentWidth >= pixelWidth)
            {
                // Make sure to keep the aspect ratio of the image
                maxWidth = min(maxWidth, parentWidth);
                double aspectRatio = (double)pixelHeight / pixelWidth;
                maxHeight = maxWidth * aspectRatio;
            }
        }

        // Prevent an image from being stretched out if it is smaller than the
        // space allocated for it (when in auto mode).
        ComPtr<IEllipse> localElementAsEllipse;
        if (SUCCEEDED(localElement.As(&localElementAsEllipse)))
        {
            // don't need to set both width and height when image size is auto since
            // we want a circle as shape.
            // max value for width should be set since adaptive cards is constrained horizontally
            RETURN_IF_FAILED(localElement->put_MaxWidth(min(maxWidth, pixelWidth)));
        }
        else
        {
            RETURN_IF_FAILED(localElement->put_MaxHeight(min(maxHeight, pixelHeight)));
            RETURN_IF_FAILED(localElement->put_MaxWidth(min(maxWidth, pixelWidth)));
        }

        if (setVisible)
        {
            ComPtr<IUIElement> frameworkElementAsUIElement;
            RETURN_IF_FAILED(localElement.As(&frameworkElementAsUIElement));
            RETURN_IF_FAILED(frameworkElementAsUIElement->put_Visibility(Visibility::Visibility_Visible));
        }

        return S_OK;
    }

    HRESULT XamlBuilder::SetMatchingHeight(_In_ IFrameworkElement* elementToChange, _In_ IFrameworkElement* elementToMatch)
    {
        DOUBLE actualHeight;
        RETURN_IF_FAILED(elementToMatch->get_ActualHeight(&actualHeight));

        ComPtr<IFrameworkElement> localElement(elementToChange);
        RETURN_IF_FAILED(localElement->put_Height(actualHeight));

        ComPtr<IUIElement> frameworkElementAsUIElement;
        RETURN_IF_FAILED(localElement.As(&frameworkElementAsUIElement));
        RETURN_IF_FAILED(frameworkElementAsUIElement->put_Visibility(Visibility::Visibility_Visible));
        return S_OK;
    }

    HRESULT XamlBuilder::HandleStylingAndPadding(_In_ IAdaptiveContainerBase* adaptiveContainer,
                                                 _In_ IBorder* containerBorder,
                                                 _In_ IAdaptiveRenderContext* renderContext,
                                                 _In_ IAdaptiveRenderArgs* renderArgs,
                                                 _Out_ ABI::AdaptiveNamespace::ContainerStyle* containerStyle)
    {
        ABI::AdaptiveNamespace::ContainerStyle localContainerStyle;
        RETURN_IF_FAILED(adaptiveContainer->get_Style(&localContainerStyle));

        ABI::AdaptiveNamespace::ContainerStyle parentContainerStyle;
        RETURN_IF_FAILED(renderArgs->get_ContainerStyle(&parentContainerStyle));

        bool hasExplicitContainerStyle{true};
        if (localContainerStyle == ABI::AdaptiveNamespace::ContainerStyle::None)
        {
            hasExplicitContainerStyle = false;
            localContainerStyle = parentContainerStyle;
        }

        ComPtr<IAdaptiveHostConfig> hostConfig;
        RETURN_IF_FAILED(renderContext->get_HostConfig(&hostConfig));

        ComPtr<IAdaptiveSpacingConfig> spacingConfig;
        RETURN_IF_FAILED(hostConfig->get_Spacing(&spacingConfig));

        UINT32 padding;
        RETURN_IF_FAILED(spacingConfig->get_Padding(&padding));
        DOUBLE paddingAsDouble = static_cast<DOUBLE>(padding);

        // If container style was explicitly assigned, apply background color and padding
        if (hasExplicitContainerStyle)
        {
            ABI::Windows::UI::Color backgroundColor;
            RETURN_IF_FAILED(GetBackgroundColorFromStyle(localContainerStyle, hostConfig.Get(), &backgroundColor));
            ComPtr<IBrush> backgroundColorBrush = XamlHelpers::GetSolidColorBrush(backgroundColor);
            RETURN_IF_FAILED(containerBorder->put_Background(backgroundColorBrush.Get()));

            // If the container style doesn't match its parent apply padding.
            if (localContainerStyle != parentContainerStyle)
            {
                Thickness paddingThickness = {paddingAsDouble, paddingAsDouble, paddingAsDouble, paddingAsDouble};
                RETURN_IF_FAILED(containerBorder->put_Padding(paddingThickness));
            }
        }

        // Find out which direction(s) we bleed in, and apply a negative margin to cause the
        // container to bleed
        ABI::AdaptiveNamespace::BleedDirection bleedDirection;
        RETURN_IF_FAILED(adaptiveContainer->get_BleedDirection(&bleedDirection));

        Thickness marginThickness = {0};
        if (bleedDirection != ABI::AdaptiveNamespace::BleedDirection::None)
        {
            if ((bleedDirection & ABI::AdaptiveNamespace::BleedDirection::Left) != ABI::AdaptiveNamespace::BleedDirection::None)
            {
                marginThickness.Left = -paddingAsDouble;
            }

            if ((bleedDirection & ABI::AdaptiveNamespace::BleedDirection::Right) != ABI::AdaptiveNamespace::BleedDirection::None)
            {
                marginThickness.Right = -paddingAsDouble;
            }

            if ((bleedDirection & ABI::AdaptiveNamespace::BleedDirection::Up) != ABI::AdaptiveNamespace::BleedDirection::None)
            {
                marginThickness.Top = -paddingAsDouble;
            }

            if ((bleedDirection & ABI::AdaptiveNamespace::BleedDirection::Down) != ABI::AdaptiveNamespace::BleedDirection::None)
            {
                marginThickness.Bottom = -paddingAsDouble;
            }

            ComPtr<IBorder> localContainerBorder(containerBorder);
            ComPtr<IFrameworkElement> containerBorderAsFrameworkElement;
            RETURN_IF_FAILED(localContainerBorder.As(&containerBorderAsFrameworkElement));
            RETURN_IF_FAILED(containerBorderAsFrameworkElement->put_Margin(marginThickness));
        }

        *containerStyle = localContainerStyle;

        return S_OK;
    }

    void XamlBuilder::AddInputValueToContext(_In_ IAdaptiveRenderContext* renderContext,
                                             _In_ IAdaptiveCardElement* adaptiveCardElement,
                                             _In_ IUIElement* inputUiElement)
    {
        ComPtr<IAdaptiveCardElement> cardElement(adaptiveCardElement);
        ComPtr<IAdaptiveInputElement> inputElement;
        THROW_IF_FAILED(cardElement.As(&inputElement));

        ComPtr<InputValue> input;
        THROW_IF_FAILED(MakeAndInitialize<InputValue>(&input, inputElement.Get(), inputUiElement));
        THROW_IF_FAILED(renderContext->AddInputValue(input.Get()));
    }

    static HRESULT HandleKeydownForInlineAction(_In_ IKeyRoutedEventArgs* args,
                                                _In_ IAdaptiveActionInvoker* actionInvoker,
                                                _In_ IAdaptiveActionElement* inlineAction)
    {
        ABI::Windows::System::VirtualKey key;
        RETURN_IF_FAILED(args->get_Key(&key));

        if (key == ABI::Windows::System::VirtualKey::VirtualKey_Enter)
        {
            ComPtr<ABI::Windows::UI::Core::ICoreWindowStatic> coreWindowStatics;
            RETURN_IF_FAILED(GetActivationFactory(HStringReference(RuntimeClass_Windows_UI_Core_CoreWindow).Get(), &coreWindowStatics));

            ComPtr<ABI::Windows::UI::Core::ICoreWindow> coreWindow;
            RETURN_IF_FAILED(coreWindowStatics->GetForCurrentThread(&coreWindow));

            ABI::Windows::UI::Core::CoreVirtualKeyStates shiftKeyState;
            RETURN_IF_FAILED(coreWindow->GetKeyState(ABI::Windows::System::VirtualKey_Shift, &shiftKeyState));

            ABI::Windows::UI::Core::CoreVirtualKeyStates ctrlKeyState;
            RETURN_IF_FAILED(coreWindow->GetKeyState(ABI::Windows::System::VirtualKey_Control, &ctrlKeyState));

            if (shiftKeyState == ABI::Windows::UI::Core::CoreVirtualKeyStates_None &&
                ctrlKeyState == ABI::Windows::UI::Core::CoreVirtualKeyStates_None)
            {
                RETURN_IF_FAILED(actionInvoker->SendActionEvent(inlineAction));
                RETURN_IF_FAILED(args->put_Handled(true));
            }
        }

        return S_OK;
    }

    static bool WarnForInlineShowCard(_In_ IAdaptiveRenderContext* renderContext,
                                      _In_ IAdaptiveActionElement* action,
                                      const std::wstring& warning)
    {
        if (action != nullptr)
        {
            ABI::AdaptiveNamespace::ActionType actionType;
            THROW_IF_FAILED(action->get_ActionType(&actionType));

            if (actionType == ABI::AdaptiveNamespace::ActionType::ShowCard)
            {
                THROW_IF_FAILED(renderContext->AddWarning(ABI::AdaptiveNamespace::WarningStatusCode::UnsupportedValue,
                                                          HStringReference(warning.c_str()).Get()));
                return true;
            }
        }

        return false;
    }

    void XamlBuilder::HandleInlineAcion(_In_ IAdaptiveRenderContext* renderContext,
                                        _In_ IAdaptiveRenderArgs* renderArgs,
                                        _In_ ITextBox* textBox,
                                        _In_ IAdaptiveActionElement* inlineAction,
                                        _COM_Outptr_ IUIElement** textBoxWithInlineAction)
    {
        ComPtr<ITextBox> localTextBox(textBox);
        ComPtr<IAdaptiveActionElement> localInlineAction(inlineAction);

        ABI::AdaptiveNamespace::ActionType actionType;
        THROW_IF_FAILED(localInlineAction->get_ActionType(&actionType));

        ComPtr<IAdaptiveHostConfig> hostConfig;
        THROW_IF_FAILED(renderContext->get_HostConfig(&hostConfig));

        // Inline ShowCards are not supported for inline actions
        if (WarnForInlineShowCard(renderContext, localInlineAction.Get(), L"Inline ShowCard not supported for InlineAction"))
        {
            THROW_IF_FAILED(localTextBox.CopyTo(textBoxWithInlineAction));
            return;
        }

        // Create a grid to hold the text box and the action button
        ComPtr<IGridStatics> gridStatics;
        THROW_IF_FAILED(GetActivationFactory(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_Grid).Get(), &gridStatics));

        ComPtr<IGrid> xamlGrid =
            XamlHelpers::CreateXamlClass<IGrid>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_Grid));
        ComPtr<IVector<ColumnDefinition*>> columnDefinitions;
        THROW_IF_FAILED(xamlGrid->get_ColumnDefinitions(&columnDefinitions));
        ComPtr<IPanel> gridAsPanel;
        THROW_IF_FAILED(xamlGrid.As(&gridAsPanel));

        // Create the first column and add the text box to it
        ComPtr<IColumnDefinition> textBoxColumnDefinition = XamlHelpers::CreateXamlClass<IColumnDefinition>(
            HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_ColumnDefinition));
        THROW_IF_FAILED(textBoxColumnDefinition->put_Width({1, GridUnitType::GridUnitType_Star}));
        THROW_IF_FAILED(columnDefinitions->Append(textBoxColumnDefinition.Get()));

        ComPtr<IFrameworkElement> textBoxAsFrameworkElement;
        THROW_IF_FAILED(localTextBox.As(&textBoxAsFrameworkElement));

        THROW_IF_FAILED(gridStatics->SetColumn(textBoxAsFrameworkElement.Get(), 0));
        XamlHelpers::AppendXamlElementToPanel(textBox, gridAsPanel.Get());

        // Create a separator column
        ComPtr<IColumnDefinition> separatorColumnDefinition = XamlHelpers::CreateXamlClass<IColumnDefinition>(
            HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_ColumnDefinition));
        THROW_IF_FAILED(separatorColumnDefinition->put_Width({1.0, GridUnitType::GridUnitType_Auto}));
        THROW_IF_FAILED(columnDefinitions->Append(separatorColumnDefinition.Get()));

        UINT spacingSize;
        THROW_IF_FAILED(GetSpacingSizeFromSpacing(hostConfig.Get(), ABI::AdaptiveNamespace::Spacing::Default, &spacingSize));

        auto separator = CreateSeparator(renderContext, spacingSize, 0, {0}, false);

        ComPtr<IFrameworkElement> separatorAsFrameworkElement;
        THROW_IF_FAILED(separator.As(&separatorAsFrameworkElement));

        THROW_IF_FAILED(gridStatics->SetColumn(separatorAsFrameworkElement.Get(), 1));
        XamlHelpers::AppendXamlElementToPanel(separator.Get(), gridAsPanel.Get());

        // Create a column for the button
        ComPtr<IColumnDefinition> inlineActionColumnDefinition = XamlHelpers::CreateXamlClass<IColumnDefinition>(
            HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_ColumnDefinition));
        THROW_IF_FAILED(inlineActionColumnDefinition->put_Width({0, GridUnitType::GridUnitType_Auto}));
        THROW_IF_FAILED(columnDefinitions->Append(inlineActionColumnDefinition.Get()));

        // Create a text box with the action title. This will be the tool tip if there's an icon
        // or the content of the button otherwise
        ComPtr<ITextBlock> titleTextBlock =
            XamlHelpers::CreateXamlClass<ITextBlock>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_TextBlock));
        HString title;
        THROW_IF_FAILED(localInlineAction->get_Title(title.GetAddressOf()));
        THROW_IF_FAILED(titleTextBlock->put_Text(title.Get()));

        HString iconUrl;
        THROW_IF_FAILED(localInlineAction->get_IconUrl(iconUrl.GetAddressOf()));
        ComPtr<IUIElement> actionUIElement;
        if (iconUrl != nullptr)
        {
            // Render the icon using the adaptive image renderer
            ComPtr<IAdaptiveElementRendererRegistration> elementRenderers;
            THROW_IF_FAILED(renderContext->get_ElementRenderers(&elementRenderers));
            ComPtr<IAdaptiveElementRenderer> imageRenderer;
            THROW_IF_FAILED(elementRenderers->Get(HStringReference(L"Image").Get(), &imageRenderer));

            ComPtr<IAdaptiveImage> adaptiveImage;
            THROW_IF_FAILED(MakeAndInitialize<AdaptiveImage>(&adaptiveImage));

            THROW_IF_FAILED(adaptiveImage->put_Url(iconUrl.Get()));

            ComPtr<IAdaptiveCardElement> adaptiveImageAsElement;
            THROW_IF_FAILED(adaptiveImage.As(&adaptiveImageAsElement));

            THROW_IF_FAILED(imageRenderer->Render(adaptiveImageAsElement.Get(), renderContext, renderArgs, &actionUIElement));

            // Add the tool tip
            ComPtr<IToolTip> toolTip =
                XamlHelpers::CreateXamlClass<IToolTip>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_ToolTip));
            ComPtr<IContentControl> toolTipAsContentControl;
            THROW_IF_FAILED(toolTip.As(&toolTipAsContentControl));
            THROW_IF_FAILED(toolTipAsContentControl->put_Content(titleTextBlock.Get()));

            ComPtr<IToolTipServiceStatics> toolTipService;
            THROW_IF_FAILED(GetActivationFactory(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_ToolTipService).Get(),
                                                 &toolTipService));

            ComPtr<IDependencyObject> actionAsDependencyObject;
            THROW_IF_FAILED(actionUIElement.As(&actionAsDependencyObject));

            THROW_IF_FAILED(toolTipService->SetToolTip(actionAsDependencyObject.Get(), toolTip.Get()));
        }
        else
        {
            // If there's no icon, just use the title text. Put it centered in a grid so it is
            // centered relative to the text box.
            ComPtr<IFrameworkElement> textBlockAsFrameworkElement;
            THROW_IF_FAILED(titleTextBlock.As(&textBlockAsFrameworkElement));
            THROW_IF_FAILED(textBlockAsFrameworkElement->put_VerticalAlignment(ABI::Windows::UI::Xaml::VerticalAlignment_Center));

            ComPtr<IGrid> titleGrid =
                XamlHelpers::CreateXamlClass<IGrid>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_Grid));
            ComPtr<IPanel> panel;
            THROW_IF_FAILED(titleGrid.As(&panel));
            XamlHelpers::AppendXamlElementToPanel(titleTextBlock.Get(), panel.Get());

            THROW_IF_FAILED(panel.As(&actionUIElement));
        }

        // Make the action the same size as the text box
        EventRegistrationToken eventToken;
        THROW_IF_FAILED(textBoxAsFrameworkElement->add_Loaded(
            Callback<IRoutedEventHandler>(
                [actionUIElement, textBoxAsFrameworkElement](IInspectable* /*sender*/, IRoutedEventArgs * /*args*/) -> HRESULT {
                    ComPtr<IFrameworkElement> actionFrameworkElement;
                    RETURN_IF_FAILED(actionUIElement.As(&actionFrameworkElement));

                    return SetMatchingHeight(actionFrameworkElement.Get(), textBoxAsFrameworkElement.Get());
                })
                .Get(),
            &eventToken));

        // Wrap the action in a button
        ComPtr<IUIElement> touchTargetUIElement;
        WrapInTouchTarget(nullptr, actionUIElement.Get(), localInlineAction.Get(), renderContext, false, L"Adaptive.Input.Text.InlineAction", &touchTargetUIElement);

        ComPtr<IFrameworkElement> touchTargetFrameworkElement;
        THROW_IF_FAILED(touchTargetUIElement.As(&touchTargetFrameworkElement));

        // Align to bottom so the icon stays with the bottom of the text box as it grows in the multiline case
        THROW_IF_FAILED(touchTargetFrameworkElement->put_VerticalAlignment(ABI::Windows::UI::Xaml::VerticalAlignment_Bottom));

        // Add the action to the column
        THROW_IF_FAILED(gridStatics->SetColumn(touchTargetFrameworkElement.Get(), 2));
        XamlHelpers::AppendXamlElementToPanel(touchTargetFrameworkElement.Get(), gridAsPanel.Get());

        // If this isn't a multiline input, enter should invoke the action
        ComPtr<IAdaptiveActionInvoker> actionInvoker;
        THROW_IF_FAILED(renderContext->get_ActionInvoker(&actionInvoker));

        boolean isMultiLine;
        THROW_IF_FAILED(textBox->get_AcceptsReturn(&isMultiLine));

        if (!isMultiLine)
        {
            ComPtr<IUIElement> textBoxAsUIElement;
            THROW_IF_FAILED(localTextBox.As(&textBoxAsUIElement));

            EventRegistrationToken keyDownEventToken;
            THROW_IF_FAILED(textBoxAsUIElement->add_KeyDown(
                Callback<IKeyEventHandler>([actionInvoker, localInlineAction](IInspectable* /*sender*/, IKeyRoutedEventArgs* args) -> HRESULT {
                    return HandleKeydownForInlineAction(args, actionInvoker.Get(), localInlineAction.Get());
                })
                    .Get(),
                &keyDownEventToken));
        }

        THROW_IF_FAILED(xamlGrid.CopyTo(textBoxWithInlineAction));
    }

    bool XamlBuilder::SupportsInteractivity(_In_ IAdaptiveHostConfig* hostConfig)
    {
        boolean supportsInteractivity;
        THROW_IF_FAILED(hostConfig->get_SupportsInteractivity(&supportsInteractivity));
        return Boolify(supportsInteractivity);
    }

    void XamlBuilder::WrapInTouchTarget(_In_ IAdaptiveCardElement* adaptiveCardElement,
                                        _In_ IUIElement* elementToWrap,
                                        _In_ IAdaptiveActionElement* action,
                                        _In_ IAdaptiveRenderContext* renderContext,
                                        bool fullWidth,
                                        const std::wstring& style,
                                        _COM_Outptr_ IUIElement** finalElement)
    {
        ComPtr<IAdaptiveHostConfig> hostConfig;
        THROW_IF_FAILED(renderContext->get_HostConfig(&hostConfig));

        if (WarnForInlineShowCard(renderContext, action, L"Inline ShowCard not supported for SelectAction"))
        {
            // Was inline show card, so don't wrap the element and just return
            ComPtr<IUIElement> localElementToWrap(elementToWrap);
            localElementToWrap.CopyTo(finalElement);
            return;
        }

        ComPtr<IButton> button =
            XamlHelpers::CreateXamlClass<IButton>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_Button));

        ComPtr<IContentControl> buttonAsContentControl;
        THROW_IF_FAILED(button.As(&buttonAsContentControl));
        THROW_IF_FAILED(buttonAsContentControl->put_Content(elementToWrap));

        ComPtr<IAdaptiveSpacingConfig> spacingConfig;
        THROW_IF_FAILED(hostConfig->get_Spacing(&spacingConfig));

        UINT32 cardPadding = 0;
        if (fullWidth)
        {
            THROW_IF_FAILED(spacingConfig->get_Padding(&cardPadding));
        }

        ComPtr<IFrameworkElement> buttonAsFrameworkElement;
        THROW_IF_FAILED(button.As(&buttonAsFrameworkElement));

        // We want the hit target to equally split the vertical space above and below the current item.
        // However, all we know is the spacing of the current item, which only applies to the spacing above.
        // We don't know what the spacing of the NEXT element will be, so we can't calculate the correct spacing
        // below. For now, we'll simply assume the bottom spacing is the same as the top. NOTE: Only apply spacings
        // (padding, margin) for adaptive card elements to avoid adding spacings to card-level selectAction.
        if (adaptiveCardElement != nullptr)
        {
            ABI::AdaptiveNamespace::Spacing elementSpacing;
            THROW_IF_FAILED(adaptiveCardElement->get_Spacing(&elementSpacing));
            UINT spacingSize;
            THROW_IF_FAILED(GetSpacingSizeFromSpacing(hostConfig.Get(), elementSpacing, &spacingSize));
            double topBottomPadding = spacingSize / 2.0;

            // For button padding, we apply the cardPadding and topBottomPadding (and then we negate these in the margin)
            ComPtr<IControl> buttonAsControl;
            THROW_IF_FAILED(button.As(&buttonAsControl));
            THROW_IF_FAILED(buttonAsControl->put_Padding({(double)cardPadding, topBottomPadding, (double)cardPadding, topBottomPadding}));

            double negativeCardMargin = cardPadding * -1.0;
            double negativeTopBottomMargin = topBottomPadding * -1.0;

            THROW_IF_FAILED(buttonAsFrameworkElement->put_Margin(
                {negativeCardMargin, negativeTopBottomMargin, negativeCardMargin, negativeTopBottomMargin}));
        }

        // Style the hit target button
        THROW_IF_FAILED(SetStyleFromResourceDictionary(renderContext, style.c_str(), buttonAsFrameworkElement.Get()));

        if (action != nullptr)
        {
            // If we have an action, use the title for the AutomationProperties.Name
            HString title;
            THROW_IF_FAILED(action->get_Title(title.GetAddressOf()));

            ComPtr<IDependencyObject> buttonAsDependencyObject;
            THROW_IF_FAILED(button.As(&buttonAsDependencyObject));

            ComPtr<IAutomationPropertiesStatics> automationPropertiesStatics;
            THROW_IF_FAILED(
                GetActivationFactory(HStringReference(RuntimeClass_Windows_UI_Xaml_Automation_AutomationProperties).Get(),
                                     &automationPropertiesStatics));

            THROW_IF_FAILED(automationPropertiesStatics->SetName(buttonAsDependencyObject.Get(), title.Get()));

            WireButtonClickToAction(button.Get(), action, renderContext);
        }

        THROW_IF_FAILED(button.CopyTo(finalElement));
    }

    void XamlBuilder::HandleSelectAction(_In_ IAdaptiveCardElement* adaptiveCardElement,
                                         _In_ IAdaptiveActionElement* selectAction,
                                         _In_ IAdaptiveRenderContext* renderContext,
                                         _In_ IUIElement* uiElement,
                                         bool supportsInteractivity,
                                         bool fullWidthTouchTarget,
                                         _COM_Outptr_ IUIElement** outUiElement)
    {
        if (selectAction != nullptr && supportsInteractivity)
        {
            WrapInTouchTarget(adaptiveCardElement, uiElement, selectAction, renderContext, fullWidthTouchTarget, L"Adaptive.SelectAction", outUiElement);
        }
        else
        {
            if (selectAction != nullptr)
            {
                renderContext->AddWarning(ABI::AdaptiveNamespace::WarningStatusCode::InteractivityNotSupported,
                                          HStringReference(L"SelectAction present, but Interactivity is not supported").Get());
            }

            ComPtr<IUIElement> localUiElement(uiElement);
            THROW_IF_FAILED(localUiElement.CopyTo(outUiElement));
        }
    }

    void XamlBuilder::WireButtonClickToAction(_In_ IButton* button, _In_ IAdaptiveActionElement* action, _In_ IAdaptiveRenderContext* renderContext)
    {
        // Note that this method currently doesn't support inline show card actions, it
        // assumes the caller won't call this method if inline show card is specified.
        ComPtr<IButton> localButton(button);
        ComPtr<IAdaptiveActionInvoker> actionInvoker;
        THROW_IF_FAILED(renderContext->get_ActionInvoker(&actionInvoker));
        ComPtr<IAdaptiveActionElement> strongAction(action);

        // Add click handler
        ComPtr<IButtonBase> buttonBase;
        THROW_IF_FAILED(localButton.As(&buttonBase));

        EventRegistrationToken clickToken;
        THROW_IF_FAILED(buttonBase->add_Click(Callback<IRoutedEventHandler>(
                                                  [strongAction, actionInvoker](IInspectable* /*sender*/, IRoutedEventArgs * /*args*/) -> HRESULT {
                                                      THROW_IF_FAILED(actionInvoker->SendActionEvent(strongAction.Get()));
                                                      return S_OK;
                                                  })
                                                  .Get(),
                                              &clickToken));
    }

    HRESULT XamlBuilder::AddHandledTappedEvent(_In_ IUIElement* uiElement)
    {
        if (uiElement == nullptr)
        {
            return E_INVALIDARG;
        }

        EventRegistrationToken clickToken;
        // Add Tap handler that sets the event as handled so that it doesn't propagate to the parent containers.
        return uiElement->add_Tapped(Callback<ITappedEventHandler>([](IInspectable* /*sender*/, ITappedRoutedEventArgs* args) -> HRESULT {
                                         return args->put_Handled(TRUE);
                                     })
                                         .Get(),
                                     &clickToken);
    }
}
