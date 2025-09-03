#!/usr/bin/env python3
"""
Example usage of the Discord bot script.
This demonstrates how to use the bot programmatically.
"""

import asyncio
import os
from discord_bot import send_message_and_kick_users, bot

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

async def main():
    """Main function to demonstrate bot usage."""
    from discord_bot import BOT_TOKEN
    
    # Check if bot token is set
    if BOT_TOKEN == "YOUR_BOT_TOKEN_HERE":
        print("Error: Please replace 'YOUR_BOT_TOKEN_HERE' with your actual Discord bot token in discord_bot.py!")
        return
    
    print("Starting Discord bot...")
    
    # Start the bot in the background
    bot_task = asyncio.create_task(bot.start(BOT_TOKEN))
    
    # Wait for the bot to be ready
    await bot.wait_until_ready()
    print("Bot is ready!")
    
    try:
        # Send messages and kick users
        print(f"Processing {len(EXAMPLE_USER_IDS)} users...")
        await send_message_and_kick_users(EXAMPLE_USER_IDS, EXAMPLE_GUILD_ID, CUSTOM_MESSAGE)
        print("Process completed!")
        
    except Exception as e:
        print(f"Error during processing: {e}")
    
    finally:
        # Close the bot
        await bot.close()

if __name__ == "__main__":
    # Run the example
    asyncio.run(main())