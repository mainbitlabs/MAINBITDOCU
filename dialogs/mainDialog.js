const config = require('../config');
const azurest = require('azure-storage');
// const nodeoutlook = require('nodejs-nodemailer-outlook');
const tableSvc = azurest.createTableService(config.storageA, config.accessK);
const azureTS = require('azure-table-storage-async');
// const moment = require('moment-timezone');

const { ChoiceFactory, ChoicePrompt, TextPrompt, WaterfallDialog} = require('botbuilder-dialogs');

const { CancelAndHelpDialog } = require('./cancelAndHelpDialog');
const { DocsDialog, DOCS_DIALOG } = require('./DOCS');
const { MailDialog, MAIL_DIALOG } = require('./MAIL');


const CHOICE_PROMPT = "CHOICE_PROMPT";
const TEXT_PROMPT = "TEXT_PROMPT";
const WATERFALL_DIALOG = "WATERFALL_DIALOG";

class MainDialog extends CancelAndHelpDialog {
    constructor(id){
        super(id || 'mainbitDialog');
        this.addDialog(new MailDialog());
        this.addDialog(new DocsDialog());
        this.addDialog(new TextPrompt(TEXT_PROMPT));
        this.addDialog(new ChoicePrompt(CHOICE_PROMPT));
        this.addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
            this.serieStep.bind(this),
            this.infoConfirmStep.bind(this),
            this.dispatcher.bind(this),
            this.choiceDialog.bind(this),
            this.finalDialog.bind(this)
        ]));
        this.initialDialogId = WATERFALL_DIALOG;
    }
    async serieStep(step){
        console.log('[mainDialog]:serieStep');
        
        await step.context.sendActivity('Recuerda que este bot tiene un tiempo limite de 10 minutos.');
        return await step.prompt(TEXT_PROMPT, `Por favor, **escribe el Número de Serie del equipo.**`);
    }
    
    async infoConfirmStep(step) {
        console.log('[mainDialog]:infoConfirmStep <<inicia>>');
        const details = step.options;
        const str = step.result;
        const serie = str.replace(/\s/g,'');
        details.serie = serie;
        const rowkey = details.serie;
    
        const query = new azurest.TableQuery().where('RowKey eq ?', rowkey);
        const result = await azureTS.queryCustomAsync(tableSvc, config.table1, query);
    
        if (result[0] == undefined) {  
            console.log('[mainDialog]:infoConfirmStep <<request fail>>', rowkey);
            
            await step.context.sendActivity(`La serie **${details.serie}** no se encontró en la base de datos, verifica la información y vuelve a intentarlo nuevamente.`); 
            return await step.endDialog();
        } else {
            console.log('[mainDialog]:infoConfirmStep <<success>>');
            console.log(result[0].RowKey._);
            
            for (const r of result) {
                details.proyecto =result[0].PartitionKey._;
                details.serie =result[0].RowKey._; 
                details.nombre =result[0].Nombre._;
                details.puesto =result[0].Puesto._; 
                details.telefono =result[0].Telefono._; 
                details.ext =result[0].Ext._;
                details.estado =result[0].Estado._;
                details.municipio =result[0].Municipio._;
                details.direccion =result[0].Domicilio._;
                details.administrativa =result[0].UnidadAdminsitrativa._;
                details.perfil =result[0].Perfil._;
                
                
                const msg=(`
                **Proyecto:** ${details.proyecto}
                \n\n **Número de Serie**: ${details.serie} 
                \n\n **Nombre**: ${details.nombre} 
                \n\n **Puesto:** ${details.puesto} 
                \n\n **Teléfono:** ${details.telefono} 
                \n\n **Extensión**: ${details.ext} 
                \n\n **Estado:** ${details.estado}  
                \n\n **Municipio:** ${details.municipio}  
                \n\n **Dirección:** ${details.direccion} 
                \n\n **Unidad de Administrativa:** ${details.administrativa} 
                \n\n **Perfil:** ${details.perfil} 
                `);
                await step.context.sendActivity(msg);
                return await step.prompt(CHOICE_PROMPT, {
                    prompt: '**¿Esta información es correcta?**',
                    choices: ChoiceFactory.toChoices(['Sí', 'No'])
                });
                
                }    
        }
    }
    
    async dispatcher(step) {
        console.log('[mainDialog]:dispatcher <<inicia>>');
        const selection = step.result.value;
        const details = step.options;
        switch (selection) {
            
            case 'Sí':
                return await step.prompt(CHOICE_PROMPT,{
                    prompt:'¿Que deseas realizar?',
                    choices: ChoiceFactory.toChoices(['Solicitar Formato de Resguardo', 'Enviar Resguardo firmado'])
                });
    
            case 'No':
                await step.context.sendActivity(`Hemos terminado por ahora.`); 
                return await step.beginDialog(MAIL_DIALOG, details);  
                
        }
    }
    
        async choiceDialog(step) {
            console.log('[mainDialog]:choiceDialog <<inicia>>');
            // console.log('result ?',step.result);
            const details = step.options;
            if (step.result === undefined) {
                return await step.endDialog();
            } else {
                details.answer = step.result.value;
                const answer = details.answer;
                if (!step.result) {
                }
                if (!answer) {
                    // exhausted attempts and no selection, start over
                    await step.context.sendActivity('Not a valid option. We\'ll restart the dialog ' +
                        'so you can try again!');
                    return await step.endDialog();
                }
                if (answer ==='Solicitar Formato de Resguardo') {
                    
                    await step.context.sendActivity(`Inicia diálogo para envío de formato.`); 
                    return await step.endDialog();  
                } 
                if (answer ==='Enviar Resguardo firmado') {
                    
                    return await step.beginDialog(DOCS_DIALOG, details);
                } 
        }
            console.log('[mainDialog]:choiceDialog<<termina>>');
            return await step.endDialog();
        }
    
        async finalDialog(step){
            console.log('[mainDialog]: finalDialog');
        return await step.endDialog();
        
        }
}
module.exports.MainDialog = MainDialog;