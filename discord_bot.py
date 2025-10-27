#!/usr/bin/env python3
"""
Discord Bot Script
Sends private messages to users and kicks them from the guild.
"""

import os
import asyncio
import discord
from discord.ext import commands
from typing import List

# Bot token (replace with your actual bot token)
BOT_TOKEN = "YOUR_BOT_TOKEN_HERE"

# Bot configuration
intents = discord.Intents.default()
intents.message_content = True
intents.members = True  # Required to kick members

bot = commands.Bot(command_prefix='!', intents=intents)

# Sample message to send to users (you can modify this)
# Use {user.mention} to mention the user in the message
SAMPLE_MESSAGE = """
Hello {user.mention}! This is a notification message.

You are receiving this message because of recent activity in our community.
Please review our community guidelines.

Thank you for your understanding.
"""

async def send_message_and_kick_users(user_ids: List[int], guild_id: int, message: str = SAMPLE_MESSAGE):
    """
    Send a private message to each user and then kick them from the guild.
    
    Args:
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

@bot.event
async def on_ready():
    """Event triggered when the bot is ready."""
    print(f"Bot is ready! Logged in as {bot.user.name}#{bot.user.discriminator}")
    print(f"Bot ID: {bot.user.id}")
    print(f"Connected to {len(bot.guilds)} guild(s):")
    for guild in bot.guilds:
        print(f"  - {guild.name} (ID: {guild.id})")

@bot.command(name='notify_and_kick')
async def notify_and_kick_command(ctx, guild_id: int, *user_ids: int):
    """
    Command to send messages and kick users.
    Usage: !notify_and_kick <guild_id> <user_id1> <user_id2> ...
    """
    if not user_ids:
        await ctx.send("Please provide at least one user ID.")
        return
    
    await ctx.send(f"Starting process for {len(user_ids)} users...")
    await send_message_and_kick_users(list(user_ids), guild_id)
    await ctx.send("Process completed!")

def main():
    """Main function to run the bot."""
    # Check if bot token is set
    if BOT_TOKEN == "YOUR_BOT_TOKEN_HERE":
        print("Error: Please replace 'YOUR_BOT_TOKEN_HERE' with your actual Discord bot token!")
        return
    
    print("Starting Discord bot...")
    
    try:
        bot.run(BOT_TOKEN)
    except discord.LoginFailure:
        print("Error: Invalid bot token!")
    except Exception as e:
        print(f"Error running bot: {e}")

if __name__ == "__main__":
    main()