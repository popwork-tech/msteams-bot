import { CardFactory, TeamsActivityHandler } from "botbuilder";

const generateWelcomeCard = () => {
  return {
    type: "AdaptiveCard",
    $schema: "https://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: "Welcome to Popwork! 🎉",
        weight: "Bolder",
        size: "Large",
      },
      {
        type: "TextBlock",
        text: "Hi ,\nThank you for installing **Popwork**! We're excited to have you onboard. 🚀",
        wrap: true,
      },
      {
        type: "TextBlock",
        text: "**What can Popwork do for you?**",
        weight: "Bolder",
        size: "Medium",
        spacing: "Medium",
      },
      {
        type: "TextBlock",
        text: "- 🤝 Drive great 1-to-1 meetings\n- 🌱 Foster regular feedback\n- 📊 Monitor team metrics",
        wrap: true,
      },
      {
        type: "TextBlock",
        text: "**How to get started:**",
        weight: "Bolder",
        size: "Medium",
        spacing: "Medium",
      },
      {
        type: "FactSet",
        facts: [
          {
            title: "📝 Sign up:",
            value: "[Sign Up Link](https://app.pop.work/sign-up)",
          },
          {
            title: "📚 Help & Documentation:",
            value: "[Help Documentation Link](https://help.pop.work/)",
          },
          {
            title: "💬 Contact Us/Support:",
            value: "[support@pop.work](mailto:support@pop.work)",
          },
        ],
      },
      {
        type: "TextBlock",
        text: "**Important Information**",
        weight: "Bolder",
        size: "Medium",
        spacing: "Medium",
      },
      {
        type: "TextBlock",
        text: "⚠️ This bot is a **notification-only bot**. It delivers updates directly to your Teams environment, but it **does not support direct conversations**.\n📌 For interactive features, please visit [the Popwork app](https://app.pop.work/).",
        wrap: true,
        spacing: "Small",
      },
    ],
    actions: [
      {
        type: "Action.OpenUrl",
        title: "Get started with Popwork",
        url: "https://app.pop.work/",
      },
    ],
  };
};

export class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        console.log(context);
        if (membersAdded[cnt].id) {
          const card = generateWelcomeCard();

          // Exemple dans un bot Microsoft Teams
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          });
          break;
        }
      }
      await next();
    });
  }
}
