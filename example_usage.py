#!/usr/bin/env python3
"""
Example usage of the Discord bot script.
This demonstrates how to use the bot programmatically.
"""

import asyncio
import discord
from discord.ext import commands

# Example user IDs (replace with actual Discord user IDs)
EXAMPLE_USER_IDS = [
    123456789012345678,  # Replace with actual user ID
    987654321098765432,  # Replace with actual user ID
    111222333444555666,  # Replace with actual user ID
]

# Example guild ID (replace with actual guild ID)
EXAMPLE_GUILD_ID = 123456789012345678  # Replace with actual guild ID

# Custom message (you can modify this)
# Use {user.mention} to mention the user in the message
CUSTOM_MESSAGE = """
Hello {user.mention}!

This is a custom notification message.
You are receiving this because of recent activity in our community.

Please review our community guidelines and ensure you follow them in the future.

Thank you for your understanding.
"""

async def send_message_and_kick_users(bot, user_ids, guild_id, message):
    """
    Send a private message to each user and then kick them from the guild.
    
    Args:
        bot: The Discord bot instance
        user_ids: List of Discord user IDs
        guild_id: ID of the guild to kick users from
        message: Message to send to users before kicking
    """
    guild = bot.get_guild(guild_id)
    if not guild:
        print(f"Error: Could not find guild with ID {guild_id}")
        return
    
    print(f"Starting process for {len(user_ids)} users in guild: {guild.name}")
    
    for user_id in user_ids:
        try:
            # Get the user
            user = await bot.fetch_user(user_id)
            print(f"Processing user: {user.name}#{user.discriminator} (ID: {user_id})")
            
            # Send private message
            try:
                # Format the message with user mention
                formatted_message = message.format(user=user)
                await user.send(formatted_message)
                print(f"✓ Sent message to {user.name}")
            except discord.Forbidden:
                print(f"✗ Could not send message to {user.name} (DMs disabled)")
            except Exception as e:
                print(f"✗ Error sending message to {user.name}: {e}")
            
            # Kick the user from the guild
            try:
                member = guild.get_member(user_id)
                if member:
                    await member.kick(reason="Automated kick after notification")
                    print(f"✓ Kicked {user.name} from guild")
                else:
                    print(f"✗ User {user.name} is not a member of the guild")
            except discord.Forbidden:
                print(f"✗ Cannot kick {user.name} (insufficient permissions)")
            except Exception as e:
                print(f"✗ Error kicking {user.name}: {e}")
            
            # Small delay to avoid rate limiting
            await asyncio.sleep(1)
            
        except discord.NotFound:
            print(f"✗ User with ID {user_id} not found")
        except Exception as e:
            print(f"✗ Error processing user {user_id}: {e}")

async def main():
    """Main function to demonstrate bot usage."""
    # Import the bot token from discord_bot
    from discord_bot import BOT_TOKEN
    
    # Check if bot token is set
    if BOT_TOKEN == "YOUR_BOT_TOKEN_HERE":
        print("Error: Please replace 'YOUR_BOT_TOKEN_HERE' with your actual Discord bot token in discord_bot.py!")
        return
    
    print("Starting Discord bot...")
    
    # Create bot instance with proper intents
    intents = discord.Intents.default()
    intents.message_content = True
    intents.members = True  # Required to kick members
    
    bot = commands.Bot(command_prefix='!', intents=intents)
    
    @bot.event
    async def on_ready():
        """Event triggered when the bot is ready."""
        print(f"Bot is ready! Logged in as {bot.user.name}#{bot.user.discriminator}")
        print(f"Bot ID: {bot.user.id}")
        print(f"Connected to {len(bot.guilds)} guild(s):")
        for guild in bot.guilds:
            print(f"  - {guild.name} (ID: {guild.id})")
        
        try:
            # Send messages and kick users
            print(f"Processing {len(EXAMPLE_USER_IDS)} users...")
            await send_message_and_kick_users(bot, EXAMPLE_USER_IDS, EXAMPLE_GUILD_ID, CUSTOM_MESSAGE)
            print("Process completed!")
            
        except Exception as e:
            print(f"Error during processing: {e}")
        
        finally:
            # Close the bot after processing
            await bot.close()
    
    try:
        await bot.start(BOT_TOKEN)
    except discord.LoginFailure:
        print("Error: Invalid bot token!")
    except Exception as e:
        print(f"Error running bot: {e}")

if __name__ == "__main__":
    # Run the example
    asyncio.run(main())