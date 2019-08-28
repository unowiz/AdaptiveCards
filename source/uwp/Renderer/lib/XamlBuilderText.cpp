// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
#include "pch.h"

#include "XamlBuilder.h"
#include "XamlHelpers.h"
#include "TextHelpers.h"

using namespace Microsoft::WRL;
using namespace Microsoft::WRL::Wrappers;
using namespace ABI::AdaptiveNamespace;
using namespace ABI::Windows::Foundation::Collections;
using namespace ABI::Windows::UI::Xaml;
using namespace ABI::Windows::UI::Xaml::Documents;
using namespace ABI::Windows::UI::Xaml::Controls;

namespace AdaptiveNamespace
{
    HRESULT XamlBuilder::BuildRichTextBlock(ABI::AdaptiveNamespace::IAdaptiveCardElement* adaptiveCardElement,
                                            ABI::AdaptiveNamespace::IAdaptiveRenderContext* renderContext,
                                            ABI::AdaptiveNamespace::IAdaptiveRenderArgs* renderArgs,
                                            ABI::Windows::UI::Xaml::IUIElement** textBlockControl)
    {
        // Create the top level rich text block and set it's properties
        ComPtr<IRichTextBlock> xamlRichTextBlock =
            XamlHelpers::CreateXamlClass<IRichTextBlock>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_RichTextBlock));

        ComPtr<IAdaptiveCardElement> localAdaptiveCardElement(adaptiveCardElement);
        ComPtr<IAdaptiveRichTextBlock> adaptiveRichTextBlock;
        RETURN_IF_FAILED(localAdaptiveCardElement.As(&adaptiveRichTextBlock));

        // Set the horizontal Alingment
        RETURN_IF_FAILED(SetHorizontalAlignment(adaptiveRichTextBlock.Get(), xamlRichTextBlock.Get()));

        // Get the highlighters
        ComPtr<IRichTextBlock5> xamlRichTextBlock5;
        RETURN_IF_FAILED(xamlRichTextBlock.As(&xamlRichTextBlock5));

        ComPtr<IVector<TextHighlighter*>> textHighlighters;
        RETURN_IF_FAILED(xamlRichTextBlock5->get_TextHighlighters(&textHighlighters));

        // Add a paragraph for the inlines
        ComPtr<IVector<Block*>> xamlBlocks;
        RETURN_IF_FAILED(xamlRichTextBlock->get_Blocks(&xamlBlocks));

        ComPtr<IParagraph> xamlParagraph =
            XamlHelpers::CreateXamlClass<IParagraph>(HStringReference(RuntimeClass_Windows_UI_Xaml_Documents_Paragraph));

        ComPtr<IBlock> paragraphAsBlock;
        RETURN_IF_FAILED(xamlParagraph.As(&paragraphAsBlock));
        RETURN_IF_FAILED(xamlBlocks->Append(paragraphAsBlock.Get()));

        // Add the Inlines
        ComPtr<IVector<ABI::Windows::UI::Xaml::Documents::Inline*>> xamlInlines;
        RETURN_IF_FAILED(xamlParagraph->get_Inlines(&xamlInlines));

        ComPtr<IVector<IAdaptiveInline*>> adaptiveInlines;
        RETURN_IF_FAILED(adaptiveRichTextBlock->get_Inlines(&adaptiveInlines));

        UINT currentOffset = 0;
        XamlHelpers::IterateOverVector<IAdaptiveInline>(adaptiveInlines.Get(), [&](IAdaptiveInline* adaptiveInline) {
            // We only support TextRun inlines for now
            ComPtr<IAdaptiveInline> localInline(adaptiveInline);
            ComPtr<IAdaptiveTextRun> adaptiveTextRun;
            RETURN_IF_FAILED(localInline.As(&adaptiveTextRun));

            ComPtr<IAdaptiveActionElement> selectAction;
            RETURN_IF_FAILED(adaptiveTextRun->get_SelectAction(&selectAction));

            ComPtr<IAdaptiveTextElement> adaptiveTextElement;
            RETURN_IF_FAILED(localInline.As(&adaptiveTextElement));

            HString text;
            RETURN_IF_FAILED(adaptiveTextElement->get_Text(text.GetAddressOf()));

            boolean isStrikethrough{false};
            RETURN_IF_FAILED(adaptiveTextRun->get_Strikethrough(&isStrikethrough));

            boolean isItalic{false};
            RETURN_IF_FAILED(adaptiveTextRun->get_Italic(&isItalic));

            boolean isUnderline{};
            RETURN_IF_FAILED(adaptiveTextRun->get_Underline(&isUnderline));

            UINT inlineLength;
            if (selectAction != nullptr)
            {
                // If there's a select action, create a hyperlink that triggers the action
                ComPtr<ABI::Windows::UI::Xaml::Documents::IHyperlink> hyperlink =
                    XamlHelpers::CreateXamlClass<ABI::Windows::UI::Xaml::Documents::IHyperlink>(
                        HStringReference(RuntimeClass_Windows_UI_Xaml_Documents_Hyperlink));

                ComPtr<IAdaptiveActionInvoker> actionInvoker;
                renderContext->get_ActionInvoker(&actionInvoker);

                EventRegistrationToken clickToken;
                RETURN_IF_FAILED(
                    hyperlink->add_Click(Callback<ABI::Windows::Foundation::ITypedEventHandler<Hyperlink*, HyperlinkClickEventArgs*>>(
                                             [selectAction, actionInvoker](IInspectable*, IHyperlinkClickEventArgs*) -> HRESULT {
                                                 return actionInvoker->SendActionEvent(selectAction.Get());
                                             })
                                             .Get(),
                                         &clickToken));

                // Add the run text to the hyperlink's inlines
                ComPtr<ABI::Windows::UI::Xaml::Documents::ISpan> hyperlinkAsSpan;
                RETURN_IF_FAILED(hyperlink.As(&hyperlinkAsSpan));

                ComPtr<IVector<ABI::Windows::UI::Xaml::Documents::Inline*>> hyperlinkInlines;
                RETURN_IF_FAILED(hyperlinkAsSpan->get_Inlines(hyperlinkInlines.GetAddressOf()));
                RETURN_IF_FAILED(AddSingleTextInline(adaptiveTextElement.Get(),
                                                     renderContext,
                                                     renderArgs,
                                                     text.Get(),
                                                     TextRunStyleParameters(isStrikethrough, isItalic, isUnderline, true),
                                                     hyperlinkInlines.Get(),
                                                     &inlineLength));

                ComPtr<ABI::Windows::UI::Xaml::Documents::IInline> hyperlinkAsInline;
                RETURN_IF_FAILED(hyperlink.As(&hyperlinkAsInline));

                // Add the hyperlink to the paragraph's inlines
                RETURN_IF_FAILED(xamlInlines->Append(hyperlinkAsInline.Get()));
            }
            else
            {
                // Add the text to the paragraph's inlines
                RETURN_IF_FAILED(AddSingleTextInline(adaptiveTextElement.Get(),
                                                     renderContext,
                                                     renderArgs,
                                                     text.Get(),
                                                     TextRunStyleParameters(isStrikethrough, isItalic, isUnderline, false),
                                                     xamlInlines.Get(),
                                                     &inlineLength));
            }

            boolean highlight;
            RETURN_IF_FAILED(adaptiveTextRun->get_Highlight(&highlight));

            if (highlight)
            {
                ComPtr<ITextHighlighter> textHighlighter;
                RETURN_IF_FAILED(GetHighlighter(adaptiveTextElement.Get(), renderContext, renderArgs, &textHighlighter));

                ComPtr<IVector<TextRange>> ranges;
                RETURN_IF_FAILED(textHighlighter->get_Ranges(&ranges));

                TextRange textRange = {(INT32)currentOffset, (INT32)inlineLength};
                RETURN_IF_FAILED(ranges->Append(textRange));

                RETURN_IF_FAILED(textHighlighters->Append(textHighlighter.Get()));
            }

            currentOffset += inlineLength;
            return S_OK;
        });

        return xamlRichTextBlock.CopyTo(textBlockControl);
    }

    HRESULT XamlBuilder::BuildTextBlock(_In_ IAdaptiveCardElement* adaptiveCardElement,
                                        _In_ IAdaptiveRenderContext* renderContext,
                                        _In_ IAdaptiveRenderArgs* renderArgs,
                                        _COM_Outptr_ IUIElement** textBlockControl)
    {
        ComPtr<IAdaptiveCardElement> cardElement(adaptiveCardElement);
        ComPtr<IAdaptiveTextBlock> adaptiveTextBlock;
        RETURN_IF_FAILED(cardElement.As(&adaptiveTextBlock));

        ComPtr<ITextBlock> xamlTextBlock =
            XamlHelpers::CreateXamlClass<ITextBlock>(HStringReference(RuntimeClass_Windows_UI_Xaml_Controls_TextBlock));

        ComPtr<IAdaptiveTextElement> adaptiveTextElement;
        RETURN_IF_FAILED(adaptiveTextBlock.As(&adaptiveTextElement));

        RETURN_IF_FAILED(StyleXamlTextBlockProperties(adaptiveTextBlock.Get(), renderContext, renderArgs, xamlTextBlock.Get()));

        ComPtr<IVector<ABI::Windows::UI::Xaml::Documents::Inline*>> inlines;
        RETURN_IF_FAILED(xamlTextBlock->get_Inlines(&inlines));

        RETURN_IF_FAILED(SetXamlInlines(adaptiveTextElement.Get(), renderContext, renderArgs, false, inlines.Get()));

        // Ensure left edge of text is consistent regardless of font size, so both small and large fonts
        // are flush on the left edge of the card by enabling TrimSideBearings
        ComPtr<ITextBlock2> xamlTextBlock2;
        RETURN_IF_FAILED(xamlTextBlock.As(&xamlTextBlock2));
        RETURN_IF_FAILED(xamlTextBlock2->put_OpticalMarginAlignment(OpticalMarginAlignment_TrimSideBearings));

        ComPtr<IFrameworkElement> frameworkElement;
        RETURN_IF_FAILED(xamlTextBlock.As(&frameworkElement));
        RETURN_IF_FAILED(SetStyleFromResourceDictionary(renderContext, L"Adaptive.TextBlock", frameworkElement.Get()));

        return xamlTextBlock.CopyTo(textBlockControl);
    }
}
