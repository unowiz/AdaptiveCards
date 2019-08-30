// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
#pragma once

#include "ImageLoadTracker.h"
#include "IXamlBuilderListener.h"
#include "IImageLoadTrackerListener.h"
#include "InputValue.h"
#include "RenderedAdaptiveCard.h"
#include "AdaptiveRenderContext.h"

namespace AdaptiveNamespace
{
    class AdaptiveCardRenderer;

    class XamlBuilder
        : public Microsoft::WRL::RuntimeClass<Microsoft::WRL::RuntimeClassFlags<Microsoft::WRL::RuntimeClassType::WinRtClassicComMix>, Microsoft::WRL::FtmBase, AdaptiveNamespace::IImageLoadTrackerListener>
    {
        friend HRESULT Microsoft::WRL::Details::MakeAndInitialize<AdaptiveNamespace::XamlBuilder, AdaptiveNamespace::XamlBuilder>(
            AdaptiveNamespace::XamlBuilder**);

        AdaptiveRuntimeStringClass(XamlBuilder);

    public:
        // IImageLoadTrackerListener
        STDMETHODIMP AllImagesLoaded();
        STDMETHODIMP ImagesLoadingHadError();

        static HRESULT BuildXamlTreeFromAdaptiveCard(
            _In_ ABI::AdaptiveNamespace::IAdaptiveCard* adaptiveCard,
            _COM_Outptr_ ABI::Windows::UI::Xaml::IFrameworkElement** xamlTreeRoot,
            _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
            Microsoft::WRL::ComPtr<XamlBuilder> xamlBuilder,
            ABI::AdaptiveNamespace::ContainerStyle defaultContainerStyle = ABI::AdaptiveNamespace::ContainerStyle::Default);
        HRESULT AddListener(_In_ IXamlBuilderListener* listener) noexcept;
        HRESULT RemoveListener(_In_ IXamlBuilderListener* listener) noexcept;
        void SetFixedDimensions(UINT width, UINT height) noexcept;
        void SetEnableXamlImageHandling(bool enableXamlImageHandling) noexcept;

        static HRESULT BuildPanelChildren(
            _In_ ABI::Windows::Foundation::Collections::IVector<ABI::AdaptiveNamespace::IAdaptiveCardElement*>* children,
            _In_ ABI::Windows::UI::Xaml::Controls::IPanel* parentPanel,
            _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* context,
            _In_ ABI::AdaptiveNamespace::IAdaptiveRenderArgs* renderArgs,
            std::function<void(ABI::Windows::UI::Xaml::IUIElement* child)> childCreatedCallback) noexcept;

        HRESULT BuildImage(_In_ ABI::AdaptiveNamespace::IAdaptiveCardElement* adaptiveCardElement,
                           _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                           _In_ ABI::AdaptiveNamespace::IAdaptiveRenderArgs* renderArgs,
                           _COM_Outptr_ ABI::Windows::UI::Xaml::IUIElement** imageControl);

        static HRESULT BuildAction(_In_ ABI::AdaptiveNamespace::IAdaptiveActionElement* adaptiveActionElement,
                                   _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                   _In_ ABI::AdaptiveNamespace::IAdaptiveRenderArgs* renderArgs,
                                   _Outptr_ ABI::Windows::UI::Xaml::IUIElement** actionControl);

        template<typename T>
        static HRESULT TryGetResourceFromResourceDictionaries(_In_ ABI::Windows::UI::Xaml::IResourceDictionary* resourceDictionary,
                                                              std::wstring resourceName,
                                                              _COM_Outptr_result_maybenull_ T** resource) noexcept;

        static HRESULT TryInsertResourceToResourceDictionaries(_In_ ABI::Windows::UI::Xaml::IResourceDictionary* resourceDictionary,
                                                               std::wstring resourceName,
                                                               _In_ IInspectable* value);

        static HRESULT HandleColumnWidth(_In_ ABI::AdaptiveNamespace::IAdaptiveColumn* column,
                                         boolean isVisible,
                                         _In_ ABI::Windows::UI::Xaml::Controls::IColumnDefinition* columnDefinition);

        static HRESULT HandleToggleVisibilityClick(_In_ ABI::Windows::UI::Xaml::IFrameworkElement* cardFrameworkElement,
                                                   _In_ ABI::AdaptiveNamespace::IAdaptiveActionElement* action);

        static HRESULT HandleStylingAndPadding(_In_ ABI::AdaptiveNamespace::IAdaptiveContainerBase* adaptiveContainer,
                                               _In_ ABI::Windows::UI::Xaml::Controls::IBorder* containerBorder,
                                               _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                               _In_ ABI::AdaptiveNamespace::IAdaptiveRenderArgs* renderArgs,
                                               _Out_ ABI::AdaptiveNamespace::ContainerStyle* containerStyle);

        static void HandleSelectAction(_In_ ABI::AdaptiveNamespace::IAdaptiveCardElement* adaptiveCardElement,
                                       _In_ ABI::AdaptiveNamespace::IAdaptiveActionElement* selectAction,
                                       _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                       _In_ ABI::Windows::UI::Xaml::IUIElement* uiElement,
                                       bool supportsInteractivity,
                                       bool fullWidthTouchTarget,
                                       _COM_Outptr_ ABI::Windows::UI::Xaml::IUIElement** outUiElement);

        static HRESULT SetStyleFromResourceDictionary(_In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                                      std::wstring resourceName,
                                                      _In_ ABI::Windows::UI::Xaml::IFrameworkElement* frameworkElement) noexcept;

        static bool SupportsInteractivity(_In_ ABI::AdaptiveNamespace::IAdaptiveHostConfig* hostConfig);

        template<typename T>
        static void SetVerticalContentAlignmentToChildren(_In_ T* container, _In_ ABI::AdaptiveNamespace::VerticalContentAlignment verticalContentAlignment)
        {
            ComPtr<T> localContainer(container);
            ComPtr<IWholeItemsPanel> containerAsPanel;
            THROW_IF_FAILED(localContainer.As(&containerAsPanel));

            ComPtr<WholeItemsPanel> panel = PeekInnards<WholeItemsPanel>(containerAsPanel);
            panel->SetVerticalContentAlignment(verticalContentAlignment);
        }

        static HRESULT AddHandledTappedEvent(_In_ ABI::Windows::UI::Xaml::IUIElement* uiElement);

        static void ApplyBackgroundToRoot(_In_ ABI::Windows::UI::Xaml::Controls::IPanel* rootPanel,
                                          _In_ ABI::AdaptiveNamespace::IAdaptiveBackgroundImage* backgroundImage,
                                          _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                          _In_ ABI::AdaptiveNamespace::IAdaptiveRenderArgs* renderArgs);

        static HRESULT AddRenderedControl(_In_ ABI::Windows::UI::Xaml::IUIElement* newControl,
                                          _In_ ABI::AdaptiveNamespace::IAdaptiveCardElement* element,
                                          _In_ ABI::Windows::UI::Xaml::Controls::IPanel* parentPanel,
                                          _In_ ABI::Windows::UI::Xaml::IUIElement* separator,
                                          _In_ ABI::Windows::UI::Xaml::Controls::IColumnDefinition* columnDefinition,
                                          std::function<void(ABI::Windows::UI::Xaml::IUIElement* child)> childCreatedCallback);

        static HRESULT RenderFallback(_In_ ABI::AdaptiveNamespace::IAdaptiveCardElement* currentElement,
                                      _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                      _In_ ABI::AdaptiveNamespace::IAdaptiveRenderArgs* renderArgs,
                                      _COM_Outptr_ ABI::Windows::UI::Xaml::IUIElement** result);

        static HRESULT SetSeparatorVisibility(_In_ ABI::Windows::UI::Xaml::Controls::IPanel* parentPanel);

        static void GetSeparationConfigForElement(_In_ ABI::AdaptiveNamespace::IAdaptiveCardElement* element,
                                                  _In_ ABI::AdaptiveNamespace::IAdaptiveHostConfig* hostConfig,
                                                  _Out_ UINT* spacing,
                                                  _Out_ UINT* separatorThickness,
                                                  _Out_ ABI::Windows::UI::Color* separatorColor,
                                                  _Out_ bool* needsSeparator);

        static Microsoft::WRL::ComPtr<ABI::Windows::UI::Xaml::IUIElement> CreateSeparator(_In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                                                                          UINT spacing,
                                                                                          UINT separatorThickness,
                                                                                          ABI::Windows::UI::Color separatorColor,
                                                                                          bool isHorizontal = true);

        static void AddInputValueToContext(_In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                           _In_ ABI::AdaptiveNamespace::IAdaptiveCardElement* adaptiveCardElement,
                                           _In_ ABI::Windows::UI::Xaml::IUIElement* inputUiElement);

        static void HandleInlineAcion(_In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                      _In_ ABI::AdaptiveNamespace::IAdaptiveRenderArgs* renderArgs,
                                      _In_ ABI::Windows::UI::Xaml::Controls::ITextBox* textBox,
                                      _In_ ABI::AdaptiveNamespace::IAdaptiveActionElement* inlineAction,
                                      _COM_Outptr_ ABI::Windows::UI::Xaml::IUIElement** textBoxWithInlineAction);

        static void WrapInTouchTarget(_In_ ABI::AdaptiveNamespace::IAdaptiveCardElement* adaptiveCardElement,
                                      _In_ ABI::Windows::UI::Xaml::IUIElement* elementToWrap,
                                      _In_ ABI::AdaptiveNamespace::IAdaptiveActionElement* action,
                                      _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                      bool fullWidth,
                                      const std::wstring& style,
                                      _COM_Outptr_ ABI::Windows::UI::Xaml::IUIElement** finalElement);

        static inline HRESULT WarnFallbackString(_In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                                 const std::string& warning)
        {
            HString warningMsg;
            RETURN_IF_FAILED(UTF8ToHString(warning, warningMsg.GetAddressOf()));

            RETURN_IF_FAILED(renderContext->AddWarning(ABI::AdaptiveNamespace::WarningStatusCode::PerformingFallback,
                                                       warningMsg.Get()));
            return S_OK;
        }

        static inline HRESULT WarnForFallbackContentElement(_In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                                            _In_ HSTRING parentElementType,
                                                            _In_ HSTRING fallbackElementType)
        try
        {
            std::string warning = "Performing fallback for element of type \"";
            warning.append(HStringToUTF8(parentElementType));
            warning.append("\" (fallback element type \"");
            warning.append(HStringToUTF8(fallbackElementType));
            warning.append("\")");

            return WarnFallbackString(renderContext, warning);
        }
        CATCH_RETURN;

        static inline HRESULT WarnForFallbackDrop(_In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                                  _In_ HSTRING elementType)
        try
        {
            std::string warning = "Dropping element of type \"";
            warning.append(HStringToUTF8(elementType));
            warning.append("\" for fallback");

            return WarnFallbackString(renderContext, warning);
        }
        CATCH_RETURN;

    private:
        XamlBuilder();

        ImageLoadTracker m_imageLoadTracker;
        std::set<Microsoft::WRL::ComPtr<IXamlBuilderListener>> m_listeners;
        Microsoft::WRL::ComPtr<ABI::Windows::Storage::Streams::IRandomAccessStreamStatics> m_randomAccessStreamStatics;
        std::vector<Microsoft::WRL::ComPtr<ABI::Windows::Foundation::IAsyncOperationWithProgress<ABI::Windows::Storage::Streams::IInputStream*, ABI::Windows::Web::Http::HttpProgress>>> m_getStreamOperations;
        std::vector<Microsoft::WRL::ComPtr<ABI::Windows::Foundation::IAsyncOperationWithProgress<UINT64, UINT64>>> m_copyStreamOperations;
        std::vector<Microsoft::WRL::ComPtr<ABI::Windows::Foundation::IAsyncOperationWithProgress<UINT32, UINT32>>> m_writeAsyncOperations;

        UINT m_fixedWidth = 0;
        UINT m_fixedHeight = 0;
        bool m_fixedDimensions = false;
        bool m_enableXamlImageHandling = false;
        Microsoft::WRL::ComPtr<ABI::AdaptiveNamespace::IAdaptiveCardResourceResolvers> m_resourceResolvers;

        static Microsoft::WRL::ComPtr<ABI::Windows::UI::Xaml::IUIElement>
        CreateRootCardElement(_In_ ABI::AdaptiveNamespace::IAdaptiveCard* adaptiveCard,
                              _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                              _In_ ABI::AdaptiveNamespace::IAdaptiveRenderArgs* renderArgs,
                              Microsoft::WRL::ComPtr<XamlBuilder> xamlBuilder,
                              _Outptr_ ABI::Windows::UI::Xaml::Controls::IPanel** bodyElementContainer);

        template<typename T>
        void SetAutoSize(T* destination, IInspectable* parentElement, IInspectable* imageContainer, bool isVisible, bool imageFiresOpenEvent);

        template<typename T>
        void SetImageSource(T* destination,
                            ABI::Windows::UI::Xaml::Media::IImageSource* imageSource,
                            ABI::Windows::UI::Xaml::Media::Stretch stretch = Stretch_UniformToFill);
        template<typename T>
        void SetImageOnUIElement(_In_ ABI::Windows::Foundation::IUriRuntimeClass* imageUrl,
                                 T* uiElement,
                                 ABI::AdaptiveNamespace::IAdaptiveCardResourceResolvers* resolvers,
                                 bool isSizeAuto,
                                 IInspectable* parentElement,
                                 IInspectable* imageContainer,
                                 bool isVisible,
                                 _Out_ bool* mustHideElement,
                                 ABI::Windows::UI::Xaml::Media::Stretch stretch = Stretch_UniformToFill);

        template<typename T>
        void PopulateImageFromUrlAsync(_In_ ABI::Windows::Foundation::IUriRuntimeClass* imageUrl, _In_ T* imageControl);
        void FireAllImagesLoaded();
        void FireImagesLoadingHadError();

        static void ArrangeButtonContent(_In_ ABI::AdaptiveNamespace::IAdaptiveActionElement* action,
                                         _In_ ABI::AdaptiveNamespace::IAdaptiveActionsConfig* actionsConfig,
                                         _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                         ABI::AdaptiveNamespace::ContainerStyle containerStyle,
                                         _In_ ABI::AdaptiveNamespace::IAdaptiveHostConfig* hostConfig,
                                         bool allActionsHaveIcons,
                                         _In_ ABI::Windows::UI::Xaml::Controls::IButton* button);

        static HRESULT BuildActions(_In_ ABI::AdaptiveNamespace::IAdaptiveCard* adaptiveCard,
                                    _In_ ABI::Windows::Foundation::Collections::IVector<ABI::AdaptiveNamespace::IAdaptiveActionElement*>* children,
                                    _In_ ABI::Windows::UI::Xaml::Controls::IPanel* bodyPanel,
                                    bool insertSeparator,
                                    _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                    _In_ ABI::AdaptiveNamespace::IAdaptiveRenderArgs* renderArgs);

        static void AddSeparatorIfNeeded(int& currentElement,
                                         _In_ ABI::AdaptiveCards::Rendering::Uwp::IAdaptiveCardElement* element,
                                         _In_ ABI::AdaptiveCards::Rendering::Uwp::IAdaptiveHostConfig* hostConfig,
                                         _In_ ABI::AdaptiveCards::Rendering::Uwp::IAdaptiveRenderContext* renderContext,
                                         _In_ ABI::Windows::UI::Xaml::Controls::IPanel* parentPanel,
                                         _Outptr_ ABI::Windows::UI::Xaml::IUIElement** addedSeparator);

        static void WireButtonClickToAction(_In_ ABI::Windows::UI::Xaml::Controls::IButton* button,
                                            _In_ ABI::AdaptiveNamespace::IAdaptiveActionElement* action,
                                            _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext);

        static HRESULT SetAutoImageSize(_In_ ABI::Windows::UI::Xaml::IFrameworkElement* imageControl,
                                        _In_ IInspectable* parentElement,
                                        _In_ ABI::Windows::UI::Xaml::Media::Imaging::IBitmapSource* imageSource,
                                        bool setVisible);

        static HRESULT SetMatchingHeight(_In_ ABI::Windows::UI::Xaml::IFrameworkElement* elementToChange,
                                         _In_ ABI::Windows::UI::Xaml::IFrameworkElement* elementToMatch);

        static HRESULT HandleActionStyling(_In_ ABI::AdaptiveNamespace::IAdaptiveActionElement* adaptiveActionElement,
                                           _In_ ABI::Windows::UI::Xaml::IFrameworkElement* buttonFrameworkElement,
                                           _In_ ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext);
    };
}
