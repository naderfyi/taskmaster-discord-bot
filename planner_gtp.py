import asyncio
import nextcord
from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.users_request_builder import UsersRequestBuilder
from nextcord import SlashOption
from msgraph.generated.models.planner_assignments import PlannerAssignments
from msgraph.generated.models.planner_task import PlannerTask
import json

def load_config():
    try:
        with open('bot/config.json', 'r') as config_file:
            return json.load(config_file)
    except FileNotFoundError:
        print("Configuration file not found. Please check the path.")
        exit(1)
    except json.JSONDecodeError:
        print("Error parsing the configuration file. Please check its format.")
        exit(1)
    
config = load_config()

intents = nextcord.Intents.default()
intents.members = True
bot = nextcord.Client(intents=nextcord.Intents.all())

bot_offline_alert_task = None

TENANT_ID = config['azure']['tenant_id']
CLIENT_ID = config['azure']['client_id']
CLIENT_SECRET = config['azure']['client_secret']
PLAN_ID = config['azure']['plan_id']
BOT_TOKEN = config['discord']['token']
LOGS_CHANNEL_ID= config['discord']['log_channel_id']
   
# Access the mappings
discord_id_mapping = config['mappings']['discord_id_mapping']
bucket_id_to_name_mapping = config['mappings']['bucket_id_to_name_mapping']
user_id_to_name_mapping = config['mappings']['user_id_to_name_mapping']
discord_channel_mapping = config['mappings']['discord_channel_mapping']

# Invert the mapping to look up bucket IDs by Discord channel IDs.
bucket_id_by_discord_channel = {v: k for k, v in discord_channel_mapping.items()}

# Invert the mapping to look up user IDs by Discord ID.
discord_id_to_user_id_mapping = {v: k for k, v in discord_id_mapping.items()}

def split_messages(tasks, max_length=2000):
    # If the entire message is shorter than the max length, return it as is.
    full_message = "".join(tasks)
    if len(full_message) <= max_length:
        return [full_message]
    
    messages = []
    current_message = ""
    for task in tasks:
        if len(current_message) + len(task) > max_length:
            messages.append(current_message)
            current_message = task
        else:
            current_message += task
    messages.append(current_message)
    return messages

async def alert_offline():
    await asyncio.sleep(60)  
    if not bot.is_ready():  
        logs_channel = bot.get_channel(LOGS_CHANNEL_ID)
        if logs_channel:
            embed = nextcord.Embed(
                title="Shop Status", description="Shop is offline", colour=0xFF0000
            )
            await logs_channel.send(embed=embed)
        else:
            print("Logs channel not found. Please check the provided ID.")

@bot.event
async def on_disconnect():
    global bot_offline_alert_task
    bot_offline_alert_task = asyncio.create_task(alert_offline())

@bot.event
async def on_ready():
    global bot_offline_alert_task
    if bot_offline_alert_task:
        bot_offline_alert_task.cancel()  
    print(f'{bot.user.name} is Ready!')

@bot.slash_command()
async def ping(interaction: nextcord.Interaction):
    await interaction.response.send_message(f"Pong! {round(bot.latency * 1000)}ms")

async def create_client_credential():
    return ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)

async def create_graph_service_client():
    client_credential = await create_client_credential()
    return GraphServiceClient(client_credential)

########################################

async def get_users():
    app_client = await create_graph_service_client()
    query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
        select=['displayName', 'id', 'mail'],
        top=25,
        orderby=['displayName']
    )
    request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
        query_parameters=query_params
    )
    users = await app_client.users.get(request_configuration=request_config)
    return users

@bot.slash_command(description="List users from Microsoft Graph")
async def list_users(interaction: nextcord.Interaction):
    users_page = await get_users()

    if users_page and users_page.value:
        response_message = ""
        for user in users_page.value:
            discord_id = discord_id_mapping.get(user.id, 'Discord ID not available')
            user_info = (f'User: **{user.display_name}**\n'
                         f'Email: {user.mail}\n'
                         f'User ID: {user.id}\n'
                         f'Discord ID: <@{discord_id}>\n\n')
            response_message += user_info
        
        # Split the response message if it's too long for Discord
        if len(response_message) >= 2000:
            # Send the message in parts if it's too long
            for i in range(0, len(response_message), 1990):
                await interaction.followup.send(response_message[i:i+1990])
        else:
            await interaction.response.send_message(response_message)
    else:
        await interaction.response.send_message("No users were found.")
########################################

async def get_user_tasks(user_id):
    app_client = await create_graph_service_client()
    if user_id:
        try:
            tasks = await app_client.users.by_user_id(user_id).planner.tasks.get()
            return tasks
        except Exception as e:
            print(f"Error fetching tasks: {e}")
            return None
    else:
        print("User ID not provided")
        return None

@bot.slash_command(description="List a user's tasks from Microsoft Planner")
async def user_tasks(
    interaction: nextcord.Interaction,
    member: nextcord.Member = SlashOption(
        name="user",
        description="The user whose tasks to list",
        required=True
    ),
    status: str = SlashOption(
        name="status",
        description="Filter tasks by status",
        required=False,
        choices={
            "Not Started": "Not Started",
            "In Progress": "In Progress",
            "Completed": "Completed"
        }
    )
):
    await interaction.response.defer()

    discord_id = str(member.id)
    user_id = discord_id_to_user_id_mapping.get(discord_id)
    
    if user_id is None:
        await interaction.followup.send("No matching user ID found for the provided Discord ID")
        return

    tasks_page = await get_user_tasks(user_id)

    if tasks_page and tasks_page.value:
        tasks_details = []
        for task in tasks_page.value:
            task_status = 'Not Started' if task.percent_complete == 0 else 'Completed' if task.percent_complete == 100 else 'In Progress'
            if status and status != task_status:
                continue

            bucket_name = bucket_id_to_name_mapping.get(task.bucket_id, "Unknown Bucket")
            
            # Check if created_by and user are not None before accessing user.id
            if task.created_by and task.created_by.user:
                creator_id = task.created_by.user.id
            else:
                creator_id = None

            creator_name = user_id_to_name_mapping.get(creator_id, "Unknown User")
            created_date = task.created_date_time.strftime('%Y-%m-%d') if task.created_date_time else "Unknown Date"
            
            task_details = (f'Task: **{task.title}**\n'
                            f'Bucket: {bucket_name}\n'
                            f'Created By: {creator_name}\n'
                            f'Status: {task_status}\n'
                            f'Created Date: {created_date}\n\n')
            tasks_details.append(task_details)

        if not tasks_details:
            await interaction.followup.send("No tasks found with the given status.")
        else:
            # Now, split the messages if necessary
            split_messages_list = split_messages(tasks_details)
            for message in split_messages_list:
                await interaction.followup.send(message)
    else:
        await interaction.followup.send("No tasks were found for the user.")

########################################
async def get_tasks_in_bucket(bucket_id):
    app_client = await create_graph_service_client()
    tasks = await app_client.planner.buckets.by_planner_bucket_id(bucket_id).tasks.get()
    return tasks

async def get_tasks_in_channel(discord_channel_id):
    bucket_id = bucket_id_by_discord_channel.get(discord_channel_id)
    if bucket_id:
        try:
            return await get_tasks_in_bucket(bucket_id)
        except Exception as e:
            print(f"Error fetching tasks for bucket ID {bucket_id}: {e}")
            return None
    else:
        print(f"No bucket found for Discord channel ID {discord_channel_id}")
        return None

@bot.slash_command(description="List tasks for a specified channel or the current channel if none is specified.")
async def channel_tasks(
    interaction: nextcord.Interaction,
    channel: nextcord.abc.GuildChannel = nextcord.SlashOption(
        required=False,
        channel_types=[nextcord.ChannelType.text],
        description="The channel to list tasks for"
    ),
    status: str = nextcord.SlashOption(
        required=False,
        choices={"Not Started": "Not Started", "In Progress": "In Progress", "Completed": "Completed"},
        description="The status of the tasks to list"
    )
):
    await interaction.response.defer()

    target_channel_id = str(channel.id if channel else interaction.channel_id)
    tasks_page = await get_tasks_in_channel(target_channel_id)

    if tasks_page and tasks_page.value:
        tasks_details = []
        for task in tasks_page.value:
            task_status = 'Not Started' if task.percent_complete == 0 else 'Completed' if task.percent_complete == 100 else 'In Progress'
            if status and status != task_status:
                continue

            assignees = task.assignments.additional_data if task.assignments else {}
            assignee_names = [user_id_to_name_mapping.get(a_id, "Unknown User") for a_id in assignees]
            
            # Check if created_by and user are not None before accessing user.id
            if task.created_by and task.created_by.user:
                creator_id = task.created_by.user.id
            else:
                creator_id = None

            creator_name = user_id_to_name_mapping.get(creator_id, "Unknown User")
            
            created_date = task.created_date_time.strftime('%Y-%m-%d') if task.created_date_time else "Unknown Date"
            task_details = (
                f'Task: **{task.title}**\n'
                f'Created By: {creator_name}\n'
                f'Created Date: {created_date}\n'
                f'Status: {task_status}\n'
                f'Assigned to: {", ".join(assignee_names)}\n\n'
            )
            tasks_details.append(task_details)

        if not tasks_details:
            await interaction.followup.send("No tasks found for this channel with the given status.")
        else:
            split_messages_list = split_messages(tasks_details)
            for message in split_messages_list:
                await interaction.followup.send(message)
    else:
        await interaction.followup.send("No tasks were found for this channel.")

async def create_planner_task(user_id,task_title, bucket_id):
    graph_client = await create_graph_service_client()
    
    new_task = PlannerTask(
        plan_id=PLAN_ID,
        bucket_id=bucket_id,
        title=task_title,
        assignments = PlannerAssignments(
            additional_data = {
                    user_id : {
                            "@odata.type" : "#microsoft.graph.plannerAssignment",
                            "orderHint" :  " !",
                    },
            }
        ),
    )

    try:
        result = await graph_client.planner.tasks.post(new_task)
        return result
    except Exception as e:
        print(f"An error occurred: {e}")
        return None
    
@bot.slash_command()
async def create_task(interaction: nextcord.Interaction, member: nextcord.Member = SlashOption(
        name="user",
        description="The user to assign the task to",
        required=True
    ), task_title: str = SlashOption(
        name="task_title",
        description="The title of the task to create",
        required=True
    ),  bucket: nextcord.abc.GuildChannel = nextcord.SlashOption(
        required=False,
        channel_types=[nextcord.ChannelType.text],
        description="The bucket to assign the task to"
    )):
    
    await interaction.response.defer()
    
    discord_id = str(member.id)
    user_id = discord_id_to_user_id_mapping.get(discord_id)
    
    if user_id is None:
        await interaction.followup.send("No matching user ID found for the provided Discord ID")
        return
    
    target_channel_id = bucket.id if bucket else interaction.channel_id
    bucket_id = bucket_id_by_discord_channel.get(str(target_channel_id))
    
    result = await create_planner_task(user_id, task_title, bucket_id)
    if result:
        await interaction.followup.send("Task created successfully.")
    else:
        await interaction.followup.send("Failed to create task.")
        
bot.run(BOT_TOKEN)
