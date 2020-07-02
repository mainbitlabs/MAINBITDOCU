const config = require('../config');
const azurest = require('azure-storage');
const image2base64 = require('image-to-base64');
const blobService = azurest.createBlobService(config.storageA,config.accessK);
const tableSvc = azurest.createTableService(config.storageA, config.accessK);
const { ComponentDialog, WaterfallDialog, ChoicePrompt, ChoiceFactory, TextPrompt, AttachmentPrompt } = require('botbuilder-dialogs');

const DOCS_DIALOG = "DOCS_DIALOG";
const CHOICE_PROMPT = "CHOICE_PROMPT";
const TEXT_PROMPT = "TEXT_PROMPT";
const ATTACH_PROMPT = "ATTACH_PROMPT";
const WATERFALL_DIALOG = 'WATERFALL_DIALOG';

class DocsDialog extends ComponentDialog {
    constructor(){
        super(DOCS_DIALOG);
        this.addDialog(new ChoicePrompt(CHOICE_PROMPT));
        this.addDialog(new TextPrompt(TEXT_PROMPT));
        this.addDialog(new AttachmentPrompt(ATTACH_PROMPT));
        this.addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
            this.adjuntaStep.bind(this),
            this.attachStep.bind(this),
            this.dispatcherStep.bind(this),
        ]));
        this.initialDialogId = WATERFALL_DIALOG;
    }
    
    async adjuntaStep(step) {
        console.log('[DocsDialog]: adjuntaStep');
        const details = step.options;
        console.log(details);
        return await step.prompt(ATTACH_PROMPT,`Adjunta aquí tu resguardo firmado`);
    }
    
   
    async attachStep(step) {
        console.log('[DocsDialog]: attachStep');
        const details = step.options;
        console.log(step.context.activity.attachments);
        
        if (step.context.activity.attachments && step.context.activity.attachments.length > 0) {
            // The user sent an attachment and the bot should handle the incoming attachment.
            const attachment = step.context.activity.attachments[0];
            const stype = attachment.contentType.split('/');
            const ctype = stype[1];
            const url = attachment.contentUrl;
            image2base64(url)
                 .then(
                     (response) => {
                         // console.log(response); //iVBORw0KGgoAAAANSwCAIA...
                         var buffer = Buffer.from(response, 'base64');
                         const blob = new Promise ((resolve, reject) => {

                             blobService.createBlockBlobFromText(config.blobcontainer, details.serie +'_'+ details.proyecto +'_ResguardoF.'+ ctype, buffer,  function(error, result, response) {
                                 if (!error) {
                                     console.log("_Archivo subido al Blob Storage",response)
                                     resolve();
                                }       
                                else{
                                    reject(
                                        console.log('Hubo un error en Blob Storage: '+ error)
                                        );
                                        
                                    }
                                });
                                });
                                
                            }
                            )
                .catch(
                    (error) => {
                        console.log(error); //Exepection error....
                    });
                    // await blob;
                    await step.context.sendActivity(`El archivo **${details.serie}_${details.proyecto}_ResguardoF.${ctype}** se ha subido correctamente`);
                    return await step.endDialog();  
                    // return await step.prompt(CHOICE_PROMPT, {
                    //     prompt: '¿Deseas adjuntar Evidencia o Documentación?',
                    //     choices: ChoiceFactory.toChoices(['Sí','No'])
                    // });
        } else {
            // Since no attachment was received, send an attachment to the user.
            await step.context.sendActivity('Por favor envía una imagen.');
        }

    }

    async dispatcherStep(step) {
        console.log('[DocsDialog]: dispatcherStep');
        const details = step.options;
        const selection = step.result.value;
        switch (selection) {
            
            case 'Sí':
                return await step.beginDialog(DOCS_DIALOG, details);
            case 'No':
            await step.context.sendActivity('De acuerdo.');             
            // TERMINA EL DIÁLOGO
            return await step.endDialog();  
            default:
                break;
        }
    }


    

}
module.exports.DocsDialog = DocsDialog;
module.exports.DOCS_DIALOG = DOCS_DIALOG;