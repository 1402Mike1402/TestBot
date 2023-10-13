const { TeamsActivityHandler, CardFactory, TurnContext } = require("botbuilder");
const rawWelcomeCard = require("./adaptiveCards/welcome.json");
const rawLearnCard = require("./adaptiveCards/learn.json");
const cardTools = require("@microsoft/adaptivecards-tools");

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    // record the likeCount
    this.likeCountObj = { likeCount: 0 };

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Trigger command by IM text
      switch (txt) {
        case "welcome": {
          const shouldHideField = true;
          var card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
    
          // Update the Adaptive Card JSON to hide 'hiddenField'
          if (shouldHideField) {
            card.body.find((item) => item.id === 'defaultInputId').isVisible = true;
          }
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        case "learn": {
          this.likeCountObj.likeCount = 0;
          const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(this.likeCountObj);
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        /**
         * case "yourCommand": {
         *   await context.sendActivity(`Add your response here!`);
         *   break;
         * }
         */
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
      }
      await next();
    });
  //   this.onAdaptiveCardInvoke(async (context, adaptiveCardAction) => {
  //     // Handle the Adaptive Card submission
  //     if (adaptiveCardAction.type === 'Action.Submit') {
  //         const submittedData = adaptiveCardAction.data;

  //         // Process the submission data and decide whether to hide 'hiddenField'
  //         const shouldHideField = true;
  //         var card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();

  //         // Update the Adaptive Card JSON to hide 'hiddenField'
  //         if (shouldHideField) {
  //           card.body.find((item) => item.id === 'defaultInputId').isVisible = false;
  //         }

  //         // Send the updated Adaptive Card back to the user
  //         await context.updateActivity({
  //             type: 'message',
  //             id: context.activity.replyToId,
  //             attachments: [CardFactory.adaptiveCard(card)],
  //         });
  //     }
  // });
    
  }

  // Invoked when an action is taken on an Adaptive Card. The Adaptive Card sends an event to the Bot and this
  // method handles that event.
  async onAdaptiveCardInvoke(context, invokeValue) {
  
    if (invokeValue.action.verb === "qwerty") {
      // Process the submission data and decide whether to hide 'text box'
      const shouldHideField = true;
      var card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();

      // Update the Adaptive Card JSON to hide 'hiddenField'
      if (shouldHideField) {
        card.body.find((item) => item.id === 'defaultInputId').isVisible = false;
        //card.body.find((item) => item.id === 'replyButton').isVisible = false;
      }

      // Send the updated Adaptive Card back to the user
      await context.updateActivity({
          type: 'message',
          id: context.activity.replyToId,
          attachments: [CardFactory.adaptiveCard(card)],
      });
      return { statusCode: 200 };
  }
    
  }
}

module.exports.TeamsBot = TeamsBot;
