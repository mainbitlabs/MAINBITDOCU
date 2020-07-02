// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.


const { ComponentDialog, DialogSet, DialogTurnStatus, TextPrompt, WaterfallDialog } = require('botbuilder-dialogs');
const { MainDialog } = require('./mainDialog');



const MAIN_WATERFALL_DIALOG = 'mainWaterfallDialog';
const MAIN_DIALOG = 'mainDialog';


class FirstDialog extends ComponentDialog {
    constructor() {
        super('MainDialog');

       

        // Define the main dialog and its related components.
        // This is a sample "book a flight" dialog.
        this.addDialog(new TextPrompt('TextPrompt'))
            .addDialog(new MainDialog(MAIN_DIALOG))
            .addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
                // this.introStep.bind(this),
                this.actStep.bind(this),
                this.finalStep.bind(this)
            ]));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;
    }

    /**
     * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {*} dialogContext
     */
    async run(context, accessor) {
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);

        const dialogContext = await dialogSet.createContext(context);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }


    async actStep(stepContext) {
        

        return await stepContext.beginDialog("mainDialog");
    }


    async finalStep(stepContext) {
        if (stepContext.result) {
            const result = stepContext.result;

            const timeProperty = new TimexProperty(result.travelDate);
            const travelDateMsg = timeProperty.toNaturalLanguage(new Date(Date.now()));
            const msg = `I have you booked to ${ result.destination } from ${ result.origin } on ${ travelDateMsg }.`;
            await stepContext.context.sendActivity(msg);
        } else {
            await stepContext.context.sendActivity('Hemos terminado por ahora.');
            return await stepContext.endDialog();
        }
        return await stepContext.endDialog();
    }
}

module.exports.FirstDialog = FirstDialog;
