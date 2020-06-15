const { BotFrameworkAdapter, MiddlewareSet, TurnContext , ActivityTypes} = require('botbuilder');

/**
 * Author: Kushal Bhabra
 * Date : 29 May 2020
 * 
 * A middleware to check if user is idle.
 * If no response for for some time, then end conversation, clear state resources.
 * 
 * This code is inspired from V3 version created by https://github.com/User1m/botbuilder-timeout/
 * It has been converted to BotFramework v4 for NodeJs v4.
 * @class IdleTimeout
 */
class IdleTimeout extends MiddlewareSet{

    /**
     * Creates an instance of IdleTimeout.
     * @param {*} options
     * @memberof IdleTimeout
     */
    constructor(options) {
        super();
        this.options = {
            PROMPT_IF_USER_IS_ACTIVE_MSG: 'Are you there?',
            PROMPT_IF_USER_IS_ACTIVE_BUTTON_TEXT: 'Yes',
            PROMPT_IF_USER_IS_ACTIVE_TIMEOUT_IN_MS: 30000, // 30 seconds
            END_CONVERSATION_MSG: "Ending conversation since you've been inactive too long. Hope to see you soon.",
            END_CONVERSATION_TIMEOUT_IN_MS: 15000 // 15 seconds
        };
        this.options = Object.assign(this.options, options);
        this.timeoutStore = new TimeoutStore();
    }

    /**
     * This gets automatically invoked by Adapter for every incoming message from user
     *
     * @param {*} context
     * @param {*} next
     * @memberof IdleTimeout
     */
    async onTurn(context, next){
        
        var _this = this;

        const convoId = context.activity.conversation.id;
        var conversationReference = TurnContext.getConversationReference(context.activity);
        var adapter = context.adapter;

        // Initialize new conversations
        if(this.timeoutStore.getStoreReferenceFor(convoId)==null)
        {
            this.timeoutStore.storeConvoIdAndConversationReference(convoId, conversationReference);
        }
        

        // Incoming activities
        if (context.activity && context.activity.text) {

            this.clearTimeoutHandlers(convoId);
            this.resetHandlers(convoId);
        }

        
        // Outgoing activities
        context.onSendActivities(async function(ctx, activities, next2){
            activities.map((outgoingActivity, index) => {
                if (outgoingActivity.text) {
                    if (outgoingActivity.type === ActivityTypes.EndOfConversation) {
                        _this.clearTimeoutHandlers(convoId);
                        _this.timeoutStore.removeConvoFromStore(convoId);
                    }
                    if(outgoingActivity.type !== ActivityTypes.EndOfConversation && _this.timeoutStore.getPromptHandlerFor(convoId) === null)
                    {
                        _this.startPromptTimer(adapter, conversationReference );
                    }
                }
            });
            await next2();
        });
        
        // Process bot logic
        await next();

        // post bot logic
    }

    
    async startEndConversationTimer(adapter, conversationReference) {
        const _this = this;
        const convoId = conversationReference.conversation.id;
        const handler = setTimeout(async () => {

            await adapter.continueConversation(conversationReference, async turnContext => {
                await turnContext.sendActivity(_this.options.END_CONVERSATION_MSG);
                await turnContext.sendActivity({
                    type : ActivityTypes.EndOfConversation
                });
            });

           /**
            * Todo: Write code to CLEAR CONVERSATION STATE AND USER STATE HERE
            * This will free up bot's memory
            */

        }, _this.options.END_CONVERSATION_TIMEOUT_IN_MS);
        _this.timeoutStore.setEndConvoHandlerFor(convoId, handler);
    }
    async startPromptTimer(adapter, conversationReference) {
        const _this = this;
        const convoId = conversationReference.conversation.id;
        const handler = setTimeout(async () => {
            
            await adapter.continueConversation(conversationReference, async turnContext => {
                await turnContext.sendActivity(_this.options.PROMPT_IF_USER_IS_ACTIVE_MSG);
            });

            _this.startEndConversationTimer(adapter, conversationReference);
        }, _this.options.PROMPT_IF_USER_IS_ACTIVE_TIMEOUT_IN_MS);
        _this.timeoutStore.setPromptHandlerFor(convoId, handler);
    }
    async clearTimeoutHandlers(convoId) {
        if (this.timeoutStore.getPromptHandlerFor(convoId) !== null) {
            clearTimeout(this.timeoutStore.getPromptHandlerFor(convoId));
        }
        if (this.timeoutStore.getEndConvoHandlerFor(convoId) !== null) {
            clearTimeout(this.timeoutStore.getEndConvoHandlerFor(convoId));
        }
    }
    async resetHandlers(convoId) {
        this.timeoutStore.setPromptHandlerFor(convoId, null);
        this.timeoutStore.setEndConvoHandlerFor(convoId, null);
    }
}

/**
 * TimeoutStore class is used to keep track of all timers
 *
 * @class TimeoutStore
 */
class TimeoutStore {
    constructor() {
        this.store = new Map();
    }
    storeConvoIdAndConversationReference(id, conversationReference) {
        const props = { conversationReference: conversationReference, promptHandler: null, endConvoHandler: null };
        this.store.set(id, props);
    }
    setPromptHandlerFor(id, value) {
        this.store.get(id).promptHandler = value;
    }
    setEndConvoHandlerFor(id, value) {
        this.store.get(id).endConvoHandler = value;
    }
    getPromptHandlerFor(id) {
        return this.store.get(id).promptHandler;
    }
    getEndConvoHandlerFor(id) {
        return this.store.get(id).endConvoHandler;
    }
    getConversationReferenceFor(id) {
        return this.store.get(id).conversationReference;
    }
    getStoreReferenceFor(id) {
        return this.store.get(id);
    }
    removeConvoFromStore(id) {
        return this.store.delete(id);
    }
}

exports.IdleTimeout = IdleTimeout;