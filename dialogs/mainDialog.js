const config = require('../config');
const azurest = require('azure-storage');
// const nodeoutlook = require('nodejs-nodemailer-outlook');
const tableSvc = azurest.createTableService(config.storageA, config.accessK);
const azureTS = require('azure-table-storage-async');
// const moment = require('moment-timezone');

const { ChoiceFactory, ChoicePrompt, TextPrompt, WaterfallDialog} = require('botbuilder-dialogs');

const { CancelAndHelpDialog } = require('./cancelAndHelpDialog');

const CHOICE_PROMPT = "CHOICE_PROMPT";
const TEXT_PROMPT = "TEXT_PROMPT";
const WATERFALL_DIALOG = "WATERFALL_DIALOG";

class MainDialog extends CancelAndHelpDialog {
    constructor(id){
        super(id || 'mainbitDialog');

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
        step.values.serie = step.result;
        const rowkey = step.values.serie;
    
        const query = new azurest.TableQuery().where('RowKey eq ?', rowkey);
        const result = await azureTS.queryCustomAsync(tableSvc,config.table1, query);
    
        if (result[0] == undefined) {  
            console.log('[mainDialog]:infoConfirmStep <<request fail>>', rowkey);
            
            await step.context.sendActivity(`La serie **${step.values.serie}** no se encontró en la base de datos, verifica la información y vuelve a intentarlo nuevamente.`); 
            return await step.endDialog();
        } else {
            console.log('[mainDialog]:infoConfirmStep <<success>>');
            console.log(result[0].RowKey._);
            
            for (const r of result) {
                config.proyecto =result[0].PartitionKey._;
                config.serie =result[0].RowKey._; 
                config.nombre =result[0].Nombre._;
                config.puesto =result[0].Puesto._; 
                config.telefono =result[0].Telefono._; 
                config.ext =result[0].Ext._;
                config.estado =result[0].Estado._;
                config.municipio =result[0].Municipio._;
                config.direccion =result[0].Domicilio._;
                config.administrativa =result[0].UnidadAdminsitrativa._;
                config.perfil =result[0].Perfil._;
                // config.inmueble =result[0].._;
                
                const msg=(`
                **Proyecto:** ${config.proyecto}
                \n\n **Número de Serie**: ${config.serie} 
                \n\n **Nombre**: ${config.nombre} 
                \n\n **Puesto:** ${config.puesto} 
                \n\n **Teléfono:** ${config.telefono} 
                \n\n **Extensión**: ${config.ext} 
                \n\n **Estado:** ${config.estado}  
                \n\n **Municipio:** ${config.municipio}  
                \n\n **Dirección:** ${config.direccion} 
                \n\n **Unidad de Administrativa:** ${config.administrativa} 
                \n\n **Perfil:** ${config.perfil} 
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
        switch (selection) {
            
            case 'Sí':
                return await step.prompt(CHOICE_PROMPT,{
                    prompt:'¿Que deseas realizar?',
                    choices: ChoiceFactory.toChoices(['Solicitar Formato de Resguardo', 'Enviar Resguardo firmado'])
                });
    
            case 'No':
                await step.context.sendActivity(`Hemos terminado por ahora.`); 

                return await step.endDialog();     
                          
              
        }
    }
    
        async choiceDialog(step) {
            console.log('[mainDialog]:choiceDialog <<inicia>>');
            // console.log('result ?',step.result);
    
            if (step.result === undefined) {
                return await step.endDialog();
            } else {
                const answer = step.result.value;
                config.solicitud = {};
                const sol = config.solicitud;
                if (!step.result) {
                }
                if (!answer) {
                    // exhausted attempts and no selection, start over
                    await step.context.sendActivity('Not a valid option. We\'ll restart the dialog ' +
                        'so you can try again!');
                    return await step.endDialog();
                }
                if (answer ==='Solicitar Formato de Resguardo') {
                    sol.level1 = answer;
                    await step.context.sendActivity(`Inicia diálogo para envío de formato.`); 
                    return await step.endDialog();

                    
                } 
                if (answer ==='Enviar Resguardo firmado') {
                    sol.level1 = answer;
                    await step.context.sendActivity(`Inicia diálogo para envío de resguardo firmado.`); 
                    return await step.endDialog();

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