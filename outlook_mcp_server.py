import datetime
import win32com.client
from typing import Optional, Dict, Any
from mcp.server.fastmcp import FastMCP

# Initialize FastMCP server
mcp = FastMCP("outlook-assistant")

# Constants
MAX_DAYS = 30
# Email cache for storing retrieved emails by number
email_cache = {}


# Helper functions
def connect_to_outlook():
    """Connect to Outlook application using COM"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace
    except Exception as e:
        raise Exception(f"Failed to connect to Outlook: {str(e)}")


def get_folder_by_name(namespace, folder_name: str):
    """Get a specific Outlook folder by name"""
    try:
        # First check inbox subfolder
        inbox = namespace.GetDefaultFolder(6)  # 6 is the index for inbox folder

        # Check inbox subfolders first (most common)
        for folder in inbox.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder

        # Then check all folders at root level
        for folder in namespace.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder

            # Also check subfolders
            for subfolder in folder.Folders:
                if subfolder.Name.lower() == folder_name.lower():
                    return subfolder

        # If not found
        return None
    except Exception as e:
        raise Exception(f"Failed to access folder {folder_name}: {str(e)}")


def format_email(mail_item) -> Dict[str, Any]:
    """Format an Outlook mail item into a structured dictionary"""
    try:
        # Extract recipients
        recipients = []
        if mail_item.Recipients:
            for i in range(1, mail_item.Recipients.Count + 1):
                recipient = mail_item.Recipients(i)
                try:
                    recipients.append(f"{recipient.Name} <{recipient.Address}>")
                except Exception:
                    recipients.append(f"{recipient.Name}")

        # Format the email data
        email_data = {
            "id": mail_item.EntryID,
            "conversation_id": (
                mail_item.ConversationID
                if hasattr(mail_item, "ConversationID")
                else None
            ),
            "subject": mail_item.Subject,
            "sender": mail_item.SenderName,
            "sender_email": mail_item.SenderEmailAddress,
            "received_time": (
                mail_item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                if mail_item.ReceivedTime
                else None
            ),
            "recipients": recipients,
            "body": mail_item.Body,
            "has_attachments": mail_item.Attachments.Count > 0,
            "attachment_count": (
                mail_item.Attachments.Count if hasattr(mail_item, "Attachments") else 0
            ),
            "unread": mail_item.UnRead if hasattr(mail_item, "UnRead") else False,
            "importance": (
                mail_item.Importance if hasattr(mail_item, "Importance") else 1
            ),
            "categories": (
                mail_item.Categories if hasattr(mail_item, "Categories") else ""
            ),
        }
        return email_data
    except Exception as e:
        raise Exception(f"Failed to format email: {str(e)}")


def clear_email_cache():
    """Clear the email cache"""
    global email_cache
    email_cache = {}


def get_emails_from_folder(folder, days: int, search_term: Optional[str] = None):
    """Get emails from a folder with optional search filter"""
    emails_list = []

    # Calculate the date threshold
    now = datetime.datetime.now()
    threshold_date = now - datetime.timedelta(days=days)

    try:
        # Set up filtering
        folder_items = folder.Items
        folder_items.Sort("[ReceivedTime]", True)  # Sort by received time, newest first

        # If we have a search term, apply it
        if search_term:
            # Handle OR operators in search term
            search_terms = [term.strip() for term in search_term.split(" OR ")]

            # Try to create a filter for subject, sender name or body
            try:
                # Build SQL filter with OR conditions for each search term
                sql_conditions = []
                for term in search_terms:
                    sql_conditions.append(
                        f"\"urn:schemas:httpmail:subject\" LIKE '%{term}%'"
                    )
                    sql_conditions.append(
                        f"\"urn:schemas:httpmail:fromname\" LIKE '%{term}%'"
                    )
                    sql_conditions.append(
                        f"\"urn:schemas:httpmail:textdescription\" LIKE '%{term}%'"
                    )

                filter_term = "@SQL=" + " OR ".join(sql_conditions)
                folder_items = folder_items.Restrict(filter_term)
            except Exception:
                # If filtering fails, we'll do manual filtering later
                pass

        # Process emails
        count = 0
        for item in folder_items:
            try:
                if hasattr(item, "ReceivedTime") and item.ReceivedTime:
                    # Convert to naive datetime for comparison
                    received_time = item.ReceivedTime.replace(tzinfo=None)

                    # Skip emails older than our threshold
                    if received_time < threshold_date:
                        continue

                    # Manual search filter if needed
                    if (
                        search_term and folder_items == folder.Items
                    ):  # If we didn't apply filter earlier
                        # Handle OR operators in search term for manual filtering
                        search_terms = [
                            term.strip().lower() for term in search_term.split(" OR ")
                        ]

                        # Check if any of the search terms match
                        found_match = False
                        for term in search_terms:
                            if (
                                term in item.Subject.lower()
                                or term in item.SenderName.lower()
                                or term in item.Body.lower()
                            ):
                                found_match = True
                                break

                        if not found_match:
                            continue

                    # Format and add the email
                    email_data = format_email(item)
                    emails_list.append(email_data)
                    count += 1
            except Exception as e:
                print(f"Warning: Error processing email: {str(e)}")
                continue

    except Exception as e:
        print(f"Error retrieving emails: {str(e)}")

    return emails_list


# MCP Tools
@mcp.tool()
def list_folders() -> str:
    """
    List all available mail folders in Outlook

    Returns:
        A list of available mail folders
    """
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()

        result = "Available mail folders:\n\n"

        # List all root folders and their subfolders
        for folder in namespace.Folders:
            result += f"- {folder.Name}\n"

            # List subfolders
            for subfolder in folder.Folders:
                result += f"  - {subfolder.Name}\n"

                # List subfolders (one more level)
                try:
                    for subsubfolder in subfolder.Folders:
                        result += f"    - {subsubfolder.Name}\n"
                except Exception:
                    pass

        return result
    except Exception as e:
        return f"Error listing mail folders: {str(e)}"


@mcp.tool()
def list_recent_emails(days: int = 7, folder_name: Optional[str] = None) -> str:
    """
    List email titles from the specified number of days

    Args:
        days: Number of days to look back for emails (max 30)
        folder_name: Name of the folder to check (if not specified, checks the Inbox)

    Returns:
        Numbered list of email titles with sender information
    """
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        return f"Error: 'days' must be an integer between 1 and {MAX_DAYS}"

    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()

        # Get the appropriate folder
        if folder_name:
            folder = get_folder_by_name(namespace, folder_name)
            if not folder:
                return f"Error: Folder '{folder_name}' not found"
        else:
            folder = namespace.GetDefaultFolder(6)  # Default inbox

        # Clear previous cache
        clear_email_cache()

        # Get emails from folder
        emails = get_emails_from_folder(folder, days)

        # Format the output and cache emails
        folder_display = f"'{folder_name}'" if folder_name else "Inbox"
        if not emails:
            return f"No emails found in {folder_display} from the last {days} days."

        result = f"Found {len(emails)} emails in {folder_display} from the last {days} days:\n\n"

        # Cache emails and build result
        for i, email in enumerate(emails, 1):
            # Store in cache
            email_cache[i] = email

            # Format for display
            result += f"Email #{i}\n"
            result += f"Subject: {email['subject']}\n"
            result += f"From: {email['sender']} <{email['sender_email']}>\n"
            result += f"Received: {email['received_time']}\n"
            result += f"Read Status: {'Read' if not email['unread'] else 'Unread'}\n"
            result += (
                f"Has Attachments: {'Yes' if email['has_attachments'] else 'No'}\n\n"
            )

        result += "To view the full content of an email, use the get_email_by_number tool with the email number."
        return result

    except Exception as e:
        return f"Error retrieving email titles: {str(e)}"


@mcp.tool()
def search_emails(
    search_term: str, days: int = 7, folder_name: Optional[str] = None
) -> str:
    """
    Search emails by contact name or keyword within a time period

    Args:
        search_term: Name or keyword to search for
        days: Number of days to look back (max 30)
        folder_name: Name of the folder to search (if not specified, searches the Inbox)

    Returns:
        Numbered list of matching email titles
    """
    if not search_term:
        return "Error: Please provide a search term"

    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        return f"Error: 'days' must be an integer between 1 and {MAX_DAYS}"

    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()

        # Get the appropriate folder
        if folder_name:
            folder = get_folder_by_name(namespace, folder_name)
            if not folder:
                return f"Error: Folder '{folder_name}' not found"
        else:
            folder = namespace.GetDefaultFolder(6)  # Default inbox

        # Clear previous cache
        clear_email_cache()

        # Get emails matching search term
        emails = get_emails_from_folder(folder, days, search_term)

        # Format the output and cache emails
        folder_display = f"'{folder_name}'" if folder_name else "Inbox"
        if not emails:
            return f"No emails matching '{search_term}' found in {folder_display} from the last {days} days."

        result = f"Found {len(emails)} emails matching '{search_term}' in {folder_display} from the last {days} days:\n\n"

        # Cache emails and build result
        for i, email in enumerate(emails, 1):
            # Store in cache
            email_cache[i] = email

            # Format for display
            result += f"Email #{i}\n"
            result += f"Subject: {email['subject']}\n"
            result += f"From: {email['sender']} <{email['sender_email']}>\n"
            result += f"Received: {email['received_time']}\n"
            result += f"Read Status: {'Read' if not email['unread'] else 'Unread'}\n"
            result += (
                f"Has Attachments: {'Yes' if email['has_attachments'] else 'No'}\n\n"
            )

        result += "To view the full content of an email, use the get_email_by_number tool with the email number."
        return result

    except Exception as e:
        return f"Error searching emails: {str(e)}"


@mcp.tool()
def get_email_by_number(email_number: int) -> str:
    """
    Get detailed content of a specific email by its number from the last listing

    Args:
        email_number: The number of the email from the list results

    Returns:
        Full details of the specified email
    """
    try:
        if not email_cache:
            return "Error: No emails have been listed yet. Please use list_recent_emails or search_emails first."

        if email_number not in email_cache:
            return f"Error: Email #{email_number} not found in the current listing."

        email_data = email_cache[email_number]

        # Connect to Outlook to get the full email content
        _, namespace = connect_to_outlook()

        # Retrieve the specific email
        email = namespace.GetItemFromID(email_data["id"])
        if not email:
            return f"Error: Email #{email_number} could not be retrieved from Outlook."

        # Format the output
        result = f"Email #{email_number} Details:\n\n"
        result += f"Subject: {email_data['subject']}\n"
        result += f"From: {email_data['sender']} <{email_data['sender_email']}>\n"
        result += f"Received: {email_data['received_time']}\n"
        result += f"Recipients: {', '.join(email_data['recipients'])}\n"
        result += (
            f"Has Attachments: {'Yes' if email_data['has_attachments'] else 'No'}\n"
        )

        if email_data["has_attachments"]:
            result += "Attachments:\n"
            for i in range(1, email.Attachments.Count + 1):
                attachment = email.Attachments(i)
                result += f"  - {attachment.FileName}\n"

        result += "\nBody:\n"
        result += email_data["body"]

        result += "\n\nTo reply to this email, use the reply_to_email_by_number tool with this email number."

        return result

    except Exception as e:
        return f"Error retrieving email details: {str(e)}"


@mcp.tool()
def reply_to_email_by_number(email_number: int, reply_text: str) -> str:
    """
    Reply to a specific email by its number from the last listing

    Args:
        email_number: The number of the email from the list results
        reply_text: The text content for the reply

    Returns:
        Status message indicating success or failure
    """
    try:
        if not email_cache:
            return "Error: No emails have been listed yet. Please use list_recent_emails or search_emails first."

        if email_number not in email_cache:
            return f"Error: Email #{email_number} not found in the current listing."

        email_id = email_cache[email_number]["id"]

        # Connect to Outlook
        outlook, namespace = connect_to_outlook()

        # Retrieve the specific email
        email = namespace.GetItemFromID(email_id)
        if not email:
            return f"Error: Email #{email_number} could not be retrieved from Outlook."

        # Create reply
        reply = email.Reply()
        reply.Body = reply_text

        # Send the reply
        reply.Send()

        return f"Reply sent successfully to: {email.SenderName} <{email.SenderEmailAddress}>"

    except Exception as e:
        return f"Error replying to email: {str(e)}"


@mcp.tool()
def compose_email(
    recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None
) -> str:
    """
    Compose and send a new email

    Args:
        recipient_email: Email address of the recipient
        subject: Subject line of the email
        body: Main content of the email
        cc_email: Email address for CC (optional)

    Returns:
        Status message indicating success or failure
    """
    try:
        # Connect to Outlook
        outlook, _ = connect_to_outlook()

        # Create a new email
        mail = outlook.CreateItem(0)  # 0 is the value for a mail item
        mail.Subject = subject
        mail.To = recipient_email

        if cc_email:
            mail.CC = cc_email

        # Add signature to the body
        mail.Body = body

        # Send the email
        mail.Send()

        return f"Email sent successfully to: {recipient_email}"

    except Exception as e:
        return f"Error sending email: {str(e)}"


@mcp.tool()
def create_folder(folder_name: str, parent_folder: Optional[str] = None) -> str:
    """
    Create a new email folder in Outlook

    Args:
        folder_name: Name of the new folder to create
        parent_folder: Name of the parent folder (if not specified, creates in the root of the mailbox)

    Returns:
        Status message indicating success or failure
    """
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()

        # Get the parent folder
        if parent_folder:
            parent = get_folder_by_name(namespace, parent_folder)
            if not parent:
                return f"Error: Parent folder '{parent_folder}' not found"
        else:
            # Use the root mailbox folder (typically the first folder which is the mailbox)
            parent = namespace.Folders[0]  # First folder is usually the mailbox

        # Check if folder already exists
        try:
            for existing_folder in parent.Folders:
                if existing_folder.Name.lower() == folder_name.lower():
                    return f"Error: Folder '{folder_name}' already exists in '{parent.Name}'"
        except Exception:
            pass

        # Create the new folder
        parent.Folders.Add(folder_name)

        parent_display = (
            f"'{parent_folder}'" if parent_folder else f"root mailbox ('{parent.Name}')"
        )
        return f"Folder '{folder_name}' created successfully in {parent_display}"

    except Exception as e:
        return f"Error creating folder: {str(e)}"


@mcp.tool()
def move_email_to_folder(email_number: int, folder_name: str) -> str:
    """
    Move a specific email by its number to a different folder

    Args:
        email_number: The number of the email from the list results
        folder_name: Name of the destination folder

    Returns:
        Status message indicating success or failure
    """
    try:
        if not email_cache:
            return "Error: No emails have been listed yet. Please use list_recent_emails or search_emails first."

        if email_number not in email_cache:
            return f"Error: Email #{email_number} not found in the current listing."

        email_id = email_cache[email_number]["id"]

        # Connect to Outlook
        _, namespace = connect_to_outlook()

        # Retrieve the specific email
        email = namespace.GetItemFromID(email_id)
        if not email:
            return f"Error: Email #{email_number} could not be retrieved from Outlook."

        # Get the destination folder
        destination_folder = get_folder_by_name(namespace, folder_name)
        if not destination_folder:
            return f"Error: Destination folder '{folder_name}' not found"

        # Move the email
        email.Move(destination_folder)

        # Remove from cache since it's moved
        email_subject = email_cache[email_number]["subject"]
        del email_cache[email_number]

        return f"Email '{email_subject}' moved successfully to folder '{folder_name}'"

    except Exception as e:
        return f"Error moving email: {str(e)}"


@mcp.tool()
def list_rules() -> str:
    """
    List all existing Outlook rules

    Returns:
        A list of all current rules with their basic information
    """
    try:
        # Connect to Outlook
        outlook, namespace = connect_to_outlook()

        # Get the rules collection
        rules = outlook.Session.DefaultStore.GetRules()

        if rules.Count == 0:
            return "No rules found in Outlook."

        result = f"Found {rules.Count} rule(s) in Outlook:\n\n"

        for i in range(1, rules.Count + 1):
            rule = rules.Item(i)
            result += f"Rule #{i}: {rule.Name}\n"
            result += f"  Enabled: {'Yes' if rule.Enabled else 'No'}\n"
            result += f"  Execution Order: {rule.ExecutionOrder}\n"

            # Get rule type
            if hasattr(rule, "RuleType"):
                rule_types = {1: "Receive", 2: "Send"}
                result += f"  Type: {rule_types.get(rule.RuleType, 'Unknown')}\n"

            result += "\n"

        return result

    except Exception as e:
        return f"Error listing rules: {str(e)}"


@mcp.tool()
def create_rule(
    rule_name: str,
    sender_contains: Optional[str] = None,
    subject_contains: Optional[str] = None,
    move_to_folder: Optional[str] = None,
    mark_as_read: bool = False,
    delete_email: bool = False,
    forward_to: Optional[str] = None,
) -> str:
    """
    Create a new Outlook rule for incoming emails

    Args:
        rule_name: Name for the new rule
        sender_contains: Text that must be contained in sender's email/name (optional)
        subject_contains: Text that must be contained in the subject (optional)
        move_to_folder: Name of folder to move matching emails to (optional)
        mark_as_read: Whether to mark matching emails as read (default: False)
        delete_email: Whether to delete matching emails (default: False)
        forward_to: Email address to forward matching emails to (optional)

    Returns:
        Status message indicating success or failure
    """
    try:
        if not sender_contains and not subject_contains:
            return "Error: At least one condition (sender_contains or subject_contains) must be specified"

        if (
            not move_to_folder
            and not mark_as_read
            and not delete_email
            and not forward_to
        ):
            return "Error: At least one action (move_to_folder, mark_as_read, delete_email, or forward_to) must be specified"

        # Connect to Outlook
        outlook, namespace = connect_to_outlook()

        # Get the rules collection
        rules = outlook.Session.DefaultStore.GetRules()

        # Check if rule name already exists
        for i in range(1, rules.Count + 1):
            if rules.Item(i).Name.lower() == rule_name.lower():
                return f"Error: A rule named '{rule_name}' already exists"

        # Create a new rule
        rule = rules.Create(rule_name, 1)  # 1 = olRuleReceive (for incoming emails)

        # Set conditions
        conditions = rule.Conditions

        if sender_contains:
            # Condition: Sender contains specific text
            sender_condition = conditions.SenderAddress
            sender_condition.Enabled = True
            sender_condition.Address = [sender_contains]

        if subject_contains:
            # Condition: Subject contains specific text
            subject_condition = conditions.Subject
            subject_condition.Enabled = True
            subject_condition.Text = [subject_contains]

        # Set actions
        actions = rule.Actions

        if move_to_folder:
            # Action: Move to folder
            destination_folder = get_folder_by_name(namespace, move_to_folder)
            if not destination_folder:
                return f"Error: Destination folder '{move_to_folder}' not found"

            move_action = actions.MoveToFolder
            move_action.Enabled = True
            move_action.Folder = destination_folder

        if mark_as_read:
            # Action: Mark as read
            read_action = actions.MarkAsRead
            read_action.Enabled = True

        if delete_email:
            # Action: Delete
            delete_action = actions.Delete
            delete_action.Enabled = True

        if forward_to:
            # Action: Forward to email address
            forward_action = actions.Forward
            forward_action.Enabled = True
            forward_action.Recipients.Add(forward_to)

        # Enable the rule
        rule.Enabled = True

        # Save the rules
        rules.Save()

        # Build summary of what was created
        summary = f"Rule '{rule_name}' created successfully!\n\nConditions:\n"
        if sender_contains:
            summary += f"- Sender contains: '{sender_contains}'\n"
        if subject_contains:
            summary += f"- Subject contains: '{subject_contains}'\n"

        summary += "\nActions:\n"
        if move_to_folder:
            summary += f"- Move to folder: '{move_to_folder}'\n"
        if mark_as_read:
            summary += "- Mark as read\n"
        if delete_email:
            summary += "- Delete email\n"
        if forward_to:
            summary += f"- Forward to: '{forward_to}'\n"

        return summary

    except Exception as e:
        return f"Error creating rule: {str(e)}"


@mcp.tool()
def delete_rule(rule_name: str) -> str:
    """
    Delete an existing Outlook rule by name

    Args:
        rule_name: Name of the rule to delete

    Returns:
        Status message indicating success or failure
    """
    try:
        # Connect to Outlook
        outlook, namespace = connect_to_outlook()

        # Get the rules collection
        rules = outlook.Session.DefaultStore.GetRules()

        # Find and delete the rule
        for i in range(1, rules.Count + 1):
            rule = rules.Item(i)
            if rule.Name.lower() == rule_name.lower():
                rule.Delete()
                rules.Save()
                return f"Rule '{rule_name}' deleted successfully"

        return f"Error: Rule '{rule_name}' not found"

    except Exception as e:
        return f"Error deleting rule: {str(e)}"


@mcp.tool()
def enable_disable_rule(rule_name: str, enabled: bool) -> str:
    """
    Enable or disable an existing Outlook rule

    Args:
        rule_name: Name of the rule to enable/disable
        enabled: True to enable, False to disable

    Returns:
        Status message indicating success or failure
    """
    try:
        # Connect to Outlook
        outlook, namespace = connect_to_outlook()

        # Get the rules collection
        rules = outlook.Session.DefaultStore.GetRules()

        # Find and modify the rule
        for i in range(1, rules.Count + 1):
            rule = rules.Item(i)
            if rule.Name.lower() == rule_name.lower():
                rule.Enabled = enabled
                rules.Save()
                status = "enabled" if enabled else "disabled"
                return f"Rule '{rule_name}' {status} successfully"

        return f"Error: Rule '{rule_name}' not found"

    except Exception as e:
        return f"Error modifying rule: {str(e)}"


# Run the server
if __name__ == "__main__":
    print("Starting Outlook MCP Server...")
    print("Connecting to Outlook...")

    try:
        # Test Outlook connection
        outlook, namespace = connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)  # 6 is inbox
        print(
            f"Successfully connected to Outlook. Inbox has {inbox.Items.Count} items."
        )

        # Run the MCP server
        print("Starting MCP server. Press Ctrl+C to stop.")
        mcp.run()
    except Exception as e:
        print(f"Error starting server: {str(e)}")
