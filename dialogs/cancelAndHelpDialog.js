// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
var config = require('../config');
const { ComponentDialog, DialogTurnStatus } = require('botbuilder-dialogs');

/**
 * This base class watches for common phrases like "help" and "cancel" and takes action on them
 * BEFORE they reach the normal bot logic.
 */
class CancelAndHelpDialog extends ComponentDialog {

    async onContinueDialog(innerDc) {
        const result = await this.interrupt(innerDc);
       
        if (result) {
            return result;
        }
        
        return await super.onContinueDialog(innerDc);
    }

    async interrupt(innerDc) {
        if (innerDc.context.activity.text) {
            const text = innerDc.context.activity.text.toLowerCase();
            switch (text) {
                case 'que pex':
                    await innerDc.context.sendActivity('Puedo ayudarte a consultar información de tu equipo');
                    return { status: DialogTurnStatus.waiting };
                    
                case 'cancel':
                case 'cancelar':
                case 'salir':
                    
                    await innerDc.context.sendActivity('Cancelando...');
                    config = {};
                    console.log(config);
                    
                    return await innerDc.cancelAllDialogs();
                case 'help':
                case '?':
                    await innerDc.context.sendActivity('[ This is where to send sample help to the user... ]');
                    return { status: DialogTurnStatus.waiting };
            }
        } else {
            return await super.onContinueDialog(innerDc);
        }
    }
}

module.exports.CancelAndHelpDialog = CancelAndHelpDialog;
