import * as React from 'react';
import * as ReactDom from 'react-dom';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
    BaseApplicationCustomizer,
    PlaceholderContent,
    PlaceholderName
} from '@microsoft/sp-application-base';

import FloatingFeedback from '../../webparts/floatingFeedback/components/FloatingFeedback';
import { IFloatingFeedbackProps } from '../../webparts/floatingFeedback/components/IFloatingFeedbackProps';

const LOG_SOURCE: string = 'FloatingFeedbackApplicationCustomizer';

export interface IFloatingFeedbackApplicationCustomizerProperties {
    // Config properties can be passed here if needed in future
}

export default class FloatingFeedbackApplicationCustomizer
    extends BaseApplicationCustomizer<IFloatingFeedbackApplicationCustomizerProperties> {

    private _topPlaceholder: PlaceholderContent | undefined;

    @override
    public onInit(): Promise<void> {
        Log.info(LOG_SOURCE, 'Initialized FloatingFeedbackApplicationCustomizer');

        this._renderPlaceHolders();

        return Promise.resolve();
    }

    private _renderPlaceHolders(): void {
        if (!this._topPlaceholder) {
            this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
                PlaceholderName.Top,
                { onDispose: this._onDispose }
            );

            if (!this._topPlaceholder) {
                console.error('The expected placeholder (Top) was not found.');
                return;
            }

            const element: React.ReactElement<IFloatingFeedbackProps> = React.createElement(
                FloatingFeedback,
                {
                    spHttpClient: this.context.spHttpClient,
                    siteUrl: this.context.pageContext.web.absoluteUrl,
                    userDisplayName: this.context.pageContext.user.displayName,
                    userEmail: this.context.pageContext.user.email,
                    pageName: document.title, // Capture current page title
                    position: 'Bottom', // Default position
                    listName: 'Feedback', // Hardcoded list name
                    description: 'Feedback Extension', // Legacy prop
                    isDarkTheme: false,
                    environmentMessage: '',
                    hasTeamsContext: false
                }
            );

            ReactDom.render(element, this._topPlaceholder.domElement);
        }
    }

    private _onDispose(): void {
        console.log('[FloatingFeedbackApplicationCustomizer._onDispose] Disposed custom top placeholder.');
    }
}
