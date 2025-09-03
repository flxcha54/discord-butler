# Discord Bot - Message and Kick Script

This Discord bot script sends private messages to users and then kicks them from a guild.

## Features

- Uses environment variable `DISCORD_BOT_TOKEN_2` for bot token
- Sends customizable private messages to users
- Kicks users from the guild after sending the message
- Includes error handling and logging
- Supports both command-line and Discord command usage

## Setup

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Set up your Discord bot:**
   - Create a Discord application at https://discord.com/developers/applications
   - Create a bot for your application
   - Get the bot token
   - Enable the following bot permissions:
     - Send Messages
     - Kick Members
     - Read Message History
   - Enable the following intents:
     - Message Content Intent
     - Server Members Intent

3. **Set environment variable:**
   ```bash
   export DISCORD_BOT_TOKEN_2="your_bot_token_here"
   ```

4. **Invite the bot to your guild:**
   - Use the OAuth2 URL generator in your Discord application
   - Select the bot scope and required permissions
   - Use the generated URL to invite the bot to your guild

## Usage

### Running the Bot

```bash
python discord_bot.py
```

### Using Discord Commands

Once the bot is running, you can use the following command in Discord:

```
!notify_and_kick <guild_id> <user_id1> <user_id2> ...
```

Example:
```
!notify_and_kick 123456789012345678 987654321098765432 111222333444555666
```

### Modifying the Message

Edit the `SAMPLE_MESSAGE` variable in `discord_bot.py` to customize the message sent to users:

```python
SAMPLE_MESSAGE = """
Your custom message here.
You can include multiple lines.
"""
```

## Bot Permissions Required

The bot needs the following permissions in the guild:
- **Kick Members**: To kick users from the guild
- **Send Messages**: To send command responses
- **Read Message History**: To read commands

## Important Notes

- The bot will attempt to send private messages to all users in the list
- If a user has DMs disabled, the bot will log this but continue with the kick
- Users must be members of the specified guild to be kicked
- The bot includes a 1-second delay between operations to avoid rate limiting
- All operations are logged to the console for monitoring

## Error Handling

The script handles various error scenarios:
- Invalid bot token
- Users not found
- Users with DMs disabled
- Insufficient permissions to kick users
- Users not in the guild
- Network errors

## Security Considerations

- Keep your bot token secure and never share it
- Only give the bot the minimum required permissions
- Monitor the bot's activity in your guild
- Consider implementing additional authentication for the kick command