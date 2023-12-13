# Discord Planner Bot

## Introduction

This Discord bot integrates with Microsoft Graph to manage tasks and user information. It allows users to list and create tasks in Microsoft Planner, view user information, and receive notifications about the bot's status.

## Requirements

- Python 3.8 or higher
- nextcord
- azure.identity.aio
- Microsoft Graph Python SDK

To install the required libraries, run:

```bash
pip install -r requirements.txt
```

## Configuration

Create a config.json file in the bot/ directory.
Replace the placeholders with your Azure and Discord credentials.

## Installation

Clone this repository and navigate to the bot directory. Install the dependencies as mentioned in the Requirements section.

## Usage

Run the bot with:

```bash
python bot.py
```

The bot supports various slash commands such as `/ping`, `/list_users`, `/user_tasks`, `/channel_tasks`, and `/create_task`.
