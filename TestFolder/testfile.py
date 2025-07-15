from tools import (
    auth_token_context,

    # attachments
    outlookMail_add_attachment,
    outlookMail_list_attachments,
    outlookMail_get_attachment,
    outlookMail_download_attachment,
    outlookMail_delete_attachment,
    outlookMail_upload_large_attachment,

    # focusedInbox
    outlookMail_delete_inference_override,
    outlookMail_update_inference_override,
    outlookMail_list_inference_overrides,

    # mailFolder
    outlookMail_delete_folder,
    outlookMail_create_mail_folder,
    outlookMail_list_folders,
    outlookMail_permanent_delete_folder,
    outlookMail_copy_folder,
    outlookMail_move_folder,
    outlookMail_get_mail_folder,
    outlookMail_update_folder_display_name,
    outlookMail_get_folder_delta,
    outlookMail_create_child_folder,
    outlookMail_list_messages_from_folder,
    outlookMail_list_child_folders,

    # mailSearchFolder
    outlookMail_delete_mail_search_folder,
    outlookMail_create_mail_search_folder,
    outlookMail_update_mail_search_folder,
    outlookMail_get_mail_search_folder,
    outlookMail_permanent_delete_mail_search_folder,
    outlookMail_get_messages_from_folder,

    # messageRule
    outlookMail_list_inbox_rules,
    outlookMail_create_message_rule,
    outlookMail_delete_message_rule,
    outlookMail_update_message_rule,
    outlookMail_get_inbox_rule_by_id,

    # messages
    outlookMail_copy_message,
    outlookMail_reply_all,
    outlookMail_send_draft,
    outlookMail_send_reply_custom,
    outlookMail_move_message,
    outlookMail_permanent_delete,
    outlookMail_create_draft_in_folder,
    outlookMail_forward_message,
    outlookMail_create_reply_all_draft,
    outlookMail_list_messages,
    outlookMail_create_draft,
    outlookMail_create_reply_draft,
    outlookMail_delete_draft,
    outlookMail_update_draft,
    outlookMail_create_forward_draft
)

message_ids = []
folder_ids = []
#print(outlookMail_list_messages())

print("-----------------Testing messages.py -----------------\n\n")
print("Test 1.1 : outlookMail_list_messages")
try:
    result = outlookMail_list_messages()
    for i in result['value']:
            message_ids.append(i['id'])
    print(result)
except Exception as e:
    print(f"Test 1.1 failed: {e}")

print("Test 1.2 : outlookMail_create_draft")


try:
    print('\nTest 1.2.1: Minimal valid input')
    result = outlookMail_create_draft(
        subject="Quick draft",
        body_content="<p>Hello, world!</p>",
        to_recipients=["user1@example.com"]
    )
    message_ids.append(result['id'])
    print(result)
except Exception as e:
    print(f"Test 1.2.1 failed: {e}")

try:
    print('\nTest 1.2.2: cc_recipients + importance=High')
    result = outlookMail_create_draft(
        subject="CC test",
        body_content="<p>With CC</p>",
        to_recipients=["user2@example.com"],
        cc_recipients=["cc@example.com"],
        importance="High"
    )
    message_ids.append(result['id'])
    print(result)
except Exception as e:
    print(f"Test 1.2.2 failed: {e}")

try:
    print('\nTest 1.2.3: bcc_recipients only')
    result = outlookMail_create_draft(
        subject="BCC test",
        body_content="<p>With BCC</p>",
        to_recipients=["user3@example.com"],
        bcc_recipients=["bcc@example.com"]
    )
    message_ids.append(result['id'])
    print(result)
except Exception as e:
    print(f"Test 1.2.3 failed: {e}")

try:
    print('\nTest 1.2.4: reply_to only')
    result = outlookMail_create_draft(
        subject="Reply-To test",
        body_content="<p>With reply-to</p>",
        to_recipients=["user4@example.com"],
        reply_to=["replyto@example.com"]
    )
    message_ids.append(result['id'])
    print(result)
except Exception as e:
    print(f"Test 1.2.4 failed: {e}")

try:
    print('\nTest 1.2.5: categories only')
    result = outlookMail_create_draft(
        subject="Categories test",
        body_content="<p>With categories</p>",
        to_recipients=["user5@example.com"],
        categories=["ProjectX", "FollowUp"]
    )
    message_ids.append(result['id'])
    print(result)
except Exception as e:
    print(f"Test 1.2.5 failed: {e}")

try:
    print('\nTest 1.2.6: All optional fields together')
    result = outlookMail_create_draft(
        subject="All fields",
        body_content="<p>Everything included</p>",
        to_recipients=["user6@example.com"],
        cc_recipients=["cc1@example.com", "cc2@example.com"],
        bcc_recipients=["bcc1@example.com"],
        reply_to=["reply1@example.com"],
        importance="Low",
        categories=["Urgent", "Review"]
    )
    message_ids.append(result['id'])
    print(result)
except Exception as e:
    print(f"Test 1.2.6 failed: {e}")

try:
    print('\nTest 1.2.7: Empty optional lists')
    result = outlookMail_create_draft(
        subject="Empty lists",
        body_content="<p>No CC/BCC/ReplyTo</p>",
        to_recipients=["user7@example.com"],
        cc_recipients=[],
        bcc_recipients=[],
        reply_to=[],
        categories=[]
    )
    message_ids.append(result['id'])
    print(result)
except Exception as e:
    print(f"Test 1.2.7 failed: {e}")


print("Test 1.3 : outlookMail_update_draft")

try:
    print('\nTest 1.3.1: Update only subject')
    result = outlookMail_update_draft(
        message_id=message_ids[0],
        subject="Updated Subject Only"
    )
    print(result)
except Exception as e:
    print(f"Test 1.3.1 failed: {e}")

try:
    print('\nTest 1.3.2: Update only body_content')
    result = outlookMail_update_draft(
        message_id=message_ids[0],
        body_content="<p>Updated body content</p>"
    )
    print(result)
except Exception as e:
    print(f"Test 1.3.2 failed: {e}")

try:
    print('\nTest 1.3.3: Update to_recipients and importance')
    result = outlookMail_update_draft(
        message_id=message_ids[0],
        to_recipients=["newto@example.com"],
        importance="High"
    )
    print(result)
except Exception as e:
    print(f"Test 1.3.3 failed: {e}")

try:
    print('\nTest 1.3.4: Update cc and bcc recipients')
    result = outlookMail_update_draft(
        message_id=message_ids[0],
        cc_recipients=["ccnew@example.com"],
        bcc_recipients=["bccnew@example.com"]
    )
    print(result)
except Exception as e:
    print(f"Test 1.3.4 failed: {e}")

try:
    print('\nTest 1.3.5: Update reply_to and categories')
    result = outlookMail_update_draft(
        message_id=message_ids[0],
        reply_to=["replyto@example.com"],
        categories=["FollowUp", "Urgent"]
    )
    print(result)
except Exception as e:
    print(f"Test 1.3.5 failed: {e}")

try:
    print('\nTest 1.3.6: Update internet_message_id, is_read, is_read_receipt_requested')
    result = outlookMail_update_draft(
        message_id=message_ids[0],
        internet_message_id="<Custommessage_ids@domain.com>",
        is_read=True,
        is_read_receipt_requested=True
    )
    print(result)
except Exception as e:
    print(f"Test 1.3.6 failed: {e}")

try:
    print('\nTest 1.3.7: Update inference_classification')
    result = outlookMail_update_draft(
        message_id=message_ids[0],
        inference_classification="focused"
    )
    print(result)
except Exception as e:
    print(f"Test 1.3.7 failed: {e}")

try:
    print('\nTest 1.3.8: Update follow-up flag')
    flag_data = {
        "flagStatus": "flagged",
        "startDateTime": {
            "dateTime": "2025-07-15T10:00:00",
            "timeZone": "Pacific Standard Time"
        },
        "dueDateTime": {
            "dateTime": "2025-07-16T17:00:00",
            "timeZone": "Pacific Standard Time"
        }
    }
    result = outlookMail_update_draft(
        message_id=message_ids[0],
        flag=flag_data
    )
    print(result)
except Exception as e:
    print(f"Test 1.3.8 failed: {e}")

try:
    print('\nTest 1.3.9: Update from_sender and sender')
    from_sender = {
        "emailAddress": {
            "address": "owner@domain.com",
            "name": "Mailbox Owner"
        }
    }
    sender = {
        "emailAddress": {
            "address": "delegate@domain.com",
            "name": "Delegate Sender"
        }
    }
    result = outlookMail_update_draft(
        message_id=message_ids[0],
        from_sender=from_sender,
        sender=sender
    )
    print(result)
except Exception as e:
    print(f"Test 1.3.9 failed: {e}")

try:
    print('\nTest 1.3.10: Update subject, body_content, to_recipients, categories, importance')
    result = outlookMail_update_draft(
        message_id=message_ids[0],
        subject="Full update",
        body_content="<p>Updated full body</p>",
        to_recipients=["combo@example.com"],
        categories=["ProjectX", "HighPriority"],
        importance="Low"
    )
    print(result)
except Exception as e:
    print(f"Test 1.3.10 failed: {e}")


print('\nTest 1.4 outlookMail_delete_draft')
try:
    result = outlookMail_delete_draft(message_id=message_ids[0])
    del message_ids[0]
    print(result)
except Exception as e:
    print(f"Test 1.4.1 failed: {e}")


print("1")
result = outlookMail_create_mail_folder("Test Folder")
folder_ids.append(result['id'])
print(result)
print('\nTest 1.5 ')
print(outlookMail_copy_message(message_ids[1],folder_ids[0]))

print('\nTest 1.6 : outlookMail_move_message')

result = outlookMail_move_message(
    message_id=message_ids[0],
    destination_folder_id=folder_ids[0]
)
print(result)


print("\n----------------- Testing attachments.py -----------------\n")

attachment_ids = []  # store created attachment IDs

print('\nTest 1.1: outlookMail_add_attachment')
result = outlookMail_add_attachment(
    message_id=message_ids[0],
    file_path='testfile.txt',  # make sure this file exists
    attachment_name='MyTestFile.txt'
)
print(result)
attachment_ids.append(result.get('id'))


print('\nTest 1.2: outlookMail_list_attachments')
result = outlookMail_list_attachments(
    message_id=message_ids[0]
)
print(result)


print('\nTest 1.3: outlookMail_get_attachment')
result = outlookMail_get_attachment(
    message_id=message_ids[0],
    attachment_id=attachment_ids[0]
)
print(result)


print('\nTest 1.4: outlookMail_download_attachment')
save_path = 'downloaded_testfile.txt'
result = outlookMail_download_attachment(
    message_id=message_ids[0],
    attachment_id=attachment_ids[0],
    save_path=save_path
)
print(f"Downloaded file saved to: {result}")


print('\nTest 1.5: outlookMail_delete_attachment')
result = outlookMail_delete_attachment(
    message_id=message_ids[0],
    attachment_id=attachment_ids[0]
)
print(result)


print('\nTest 1.6: outlookMail_upload_large_attachment')
# Create a slightly bigger test file if needed; using small file works too
result = outlookMail_upload_large_attachment(
    message_id=message_ids[0],
    file_path='testfile.txt',
    is_inline=False
)
print(result)

print("\n----------------- Testing inference_overrides.py -----------------\n")

override_ids = []

print('\nTest 1.1: outlookMail_list_inference_overrides')
result = outlookMail_list_inference_overrides()
print(result)

# get first override ID if exists
if 'value' in result and result['value']:
    override_ids.append(result['value'][0]['id'])
else:
    print("No overrides found to update/delete")
    override_ids.append(None)  # keep list length consistent


if override_ids[0]:
    print('\nTest 1.2: outlookMail_update_inference_override')
    result = outlookMail_update_inference_override(
        override_id=override_ids[0],
        classify_as="other"
    )
    print(result)

    print('\nTest 1.3: outlookMail_delete_inference_override')
    result = outlookMail_delete_inference_override(
        override_id=override_ids[0]
    )
    print(result)
else:
    print("Skipped update & delete tests because no override_id found.")



print("\n----------------- Testing mail_folders.py -----------------\n")

folder_ids = []

print("\nTest 1.1: outlookMail_list_folders()")
result = outlookMail_list_folders()
print(result)


print("\nTest 1.2: outlookMail_create_mail_folder('Test Folder')")
result = outlookMail_create_mail_folder("Test Folder")
folder_ids.append(result.get('id'))
print(result)


if folder_ids[0]:
    print("\nTest 1.3: outlookMail_get_mail_folder")
    result = outlookMail_get_mail_folder(folder_id=folder_ids[0])
    print(result)

    print("\nTest 1.4: outlookMail_update_folder_display_name")
    result = outlookMail_update_folder_display_name(folder_id=folder_ids[0], display_name="Updated Test Folder")
    print(result)

    print("\nTest 1.5: outlookMail_list_child_folders")
    result = outlookMail_list_child_folders(folder_id=folder_ids[0])
    print(result)

    print("\nTest 1.6: outlookMail_create_child_folder")
    child_result = outlookMail_create_child_folder(parent_folder_id=folder_ids[0], display_name="Child Folder")
    child_id = child_result.get('id')
    print(child_result)

    if child_id:
        print("\nTest 1.7: outlookMail_copy_folder")
        copy_result = outlookMail_copy_folder(folder_id=child_id, destination_id=folder_ids[0])
        copy_id = copy_result.get('id')
        print(copy_result)

        if copy_id:
            print("\nTest 1.8: outlookMail_move_folder")
            move_result = outlookMail_move_folder(folder_id=copy_id, destination_id=folder_ids[0])
            print(move_result)

        print("\nTest 1.9: outlookMail_delete_folder (child folder)")
        del_result = outlookMail_delete_folder(folder_id=child_id)
        print(del_result)

    print("\nTest 1.10: outlookMail_delete_folder (main folder)")
    del_result = outlookMail_delete_folder(folder_id=folder_ids[0])
    print(del_result)

else:
    print("Skipping folder detail tests: folder_id not found.")


print("\nTest 1.11: outlookMail_get_folder_delta")
result = outlookMail_get_folder_delta()
print(result)

print("\n----------------- Testing mail_search_folder.py -----------------\n")

search_folder_id = None

parent_folder_id = folder_ids[0]
source_folder_ids = [folder_ids[0]]

print("\nTest 2.1: outlookMail_create_mail_search_folder")
create_result = outlookMail_create_mail_search_folder(
    parent_folder_id=parent_folder_id,
    display_name="My Search Folder",
    include_nested_folders=True,
    source_folder_ids=source_folder_ids,
    filter_query="contains(subject, 'digest')"
)
search_folder_id = create_result.get('id')
print(create_result)

if search_folder_id:
    print("\nTest 2.2: outlookMail_get_mail_search_folder")
    get_result = outlookMail_get_mail_search_folder(folder_id=search_folder_id)
    print(get_result)

    print("\nTest 2.3: outlookMail_update_mail_search_folder - update name & disable nested")
    update_result = outlookMail_update_mail_search_folder(
        folder_id=search_folder_id,
        displayName="Updated Search Folder",
        includeNestedFolders=False
    )
    print(update_result)

    print("\nTest 2.4: outlookMail_update_mail_search_folder - change filter query")
    update_result2 = outlookMail_update_mail_search_folder(
        folder_id=search_folder_id,
        filterQuery="contains(subject, 'report')"
    )
    print(update_result2)

    print("\nTest 2.5: outlookMail_update_mail_search_folder - add extra source folder")
    extra_source_folder_ids = source_folder_ids + [folder_ids[1]] if len(folder_ids) > 1 else source_folder_ids
    update_result3 = outlookMail_update_mail_search_folder(
        folder_id=search_folder_id,
        sourceFolderIds=extra_source_folder_ids
    )
    print(update_result3)

    print("\nTest 2.6: outlookMail_get_mail_search_folder after updates")
    get_after_update = outlookMail_get_mail_search_folder(folder_id=search_folder_id)
    print(get_after_update)

else:
    print("Skipping rest of tests: search folder creation failed.")


print("\n----------------- Testing mail_search_folder extra -----------------\n")

search_folder_id = None

parent_folder_id = folder_ids[0]

print("\nTest 3.1: Create search folder")
create_result = outlookMail_create_mail_search_folder(
    parent_folder_id=parent_folder_id,
    display_name="Test Search Folder",
    include_nested_folders=True,
    source_folder_ids=[folder_ids[0]],
    filter_query="contains(subject, 'digest')"
)
print(create_result)
search_folder_id = create_result.get('id')

if search_folder_id:
    print("\nTest 3.2: Get messages from created search folder")
    messages_result = outlookMail_get_messages_from_folder(
        folder_id=search_folder_id,
        top=5,
        select="subject,sender,receivedDateTime"
    )
    print(messages_result)

    print("\nTest 3.3: Soft delete search folder")
    delete_result = outlookMail_delete_mail_search_folder(search_folder_id)
    print(delete_result)

    print("\nTest 3.4: Permanently delete search folder")
    permanent_delete_result = outlookMail_permanent_delete_mail_search_folder(search_folder_id)
    print(permanent_delete_result)

else:
    print("Skipping rest of tests: search folder creation failed.")


print("\n----------------- Testing inbox rules -----------------\n")

print("\nTest 4.1: List all inbox rules")
rules_result = outlookMail_list_inbox_rules()
print(rules_result)

rule_id = None
if "value" in rules_result and rules_result["value"]:
    rule_id = rules_result["value"][0].get("id")

if rule_id:
    print(f"\nTest 4.2: Get inbox rule by ID: {rule_id}")
    rule_detail = outlookMail_get_inbox_rule_by_id(rule_id)
    print(rule_detail)
else:
    print("No inbox rules found; skipping get_inbox_rule_by_id test.")


print("\n----------------- Testing outlookMail_create_message_rule -----------------\n")


dummy_recipient = [{
    "emailAddress": {
        "address": "testuser@example.com",
        "name": "Test User"
    }
}]

# Test 1
print("Test 1: Basic rule with moveToFolder action")
res1 = outlookMail_create_message_rule(
    displayName="Move all to folder",
    sequence=1,
    actions={"moveToFolder": folder_ids[0]}
)
print(res1)

# Test 2
print("\nTest 2: Move to folder + markAsRead")
res2 = outlookMail_create_message_rule(
    displayName="Move and mark as read",
    sequence=2,
    actions={"moveToFolder": folder_ids[1], "markAsRead": True}
)
print(res2)

# Test 3
print("\nTest 3: ForwardTo recipient")
res3 = outlookMail_create_message_rule(
    displayName="Forward emails",
    sequence=3,
    actions={"forwardTo": dummy_recipient}
)
print(res3)

# Test 4
print("\nTest 4: Move mails from testuser@example.com")
res4 = outlookMail_create_message_rule(
    displayName="Filter by sender",
    sequence=4,
    actions={"moveToFolder": folder_ids[2]},
    conditions={
        "fromAddresses": dummy_recipient
    }
)
print(res4)

# Test 5
print("\nTest 5: Subject contains 'urgent'")
res5 = outlookMail_create_message_rule(
    displayName="Subject urgent",
    sequence=5,
    actions={"moveToFolder": folder_ids[3]},
    conditions={
        "subjectContains": ["urgent"]
    }
)
print(res5)

# Test 6
print("\nTest 6: Exception for boss@example.com")
res6 = outlookMail_create_message_rule(
    displayName="Except boss",
    sequence=6,
    actions={"moveToFolder": folder_ids[2]},
    conditions={"bodyContains": ["report"]},
    exceptions={
        "fromAddresses": [{
            "emailAddress": {"address": "boss@example.com"}
        }]
    }
)
print(res6)

# Test 7
print("\nTest 7: Disabled rule initially")
res7 = outlookMail_create_message_rule(
    displayName="Disabled rule",
    sequence=7,
    actions={"moveToFolder": folder_ids[0]},
    isEnabled=False
)
print(res7)

# Test 8
print("\nTest 8: Mark importance high")
res8 = outlookMail_create_message_rule(
    displayName="Set importance",
    sequence=8,
    actions={"markImportance": "high"}
)
print(res8)

# Test 9
print("\nTest 9: Stop processing after this rule")
res9 = outlookMail_create_message_rule(
    displayName="Stop processing",
    sequence=9,
    actions={
        "moveToFolder": folder_ids[1],
        "stopProcessingRules": True
    }
)
print(res9)

# Test 10
print("\nTest 10: Move, mark as read, assign category")
res10 = outlookMail_create_message_rule(
    displayName="Multi-action rule",
    sequence=10,
    actions={
        "moveToFolder": folder_ids[2],
        "markAsRead": True,
        "assignCategories": ["Reports"]
    }
)
print(res10)

print("\n----- Robust Test: outlookMail_update_message_rule (using only rule_ids[0-2] & folders_id[0-2]) -----\n")

response = outlookMail_list_inbox_rules()
if "value" in response:
    rule_ids = [rule["id"] for rule in response["value"]]
else:
    rule_ids = []


dummy_recipient = [{
    "emailAddress": {"address": "updateuser@example.com", "name": "Update User"}
}]

# Test 1: Rename only
print("\nTest 1: Rename rule_id=0")
res1 = outlookMail_update_message_rule(
    rule_id=rule_ids[0],
    displayName="Rule Rename Test"
)
print(res1)

# Test 2: Change sequence only
print("\nTest 2: Change sequence to 99 for rule_id=1")
res2 = outlookMail_update_message_rule(
    rule_id=rule_ids[1],
    sequence=99
)
print(res2)

# Test 3: Disable rule
print("\nTest 3: Disable rule_id=2")
res3 = outlookMail_update_message_rule(
    rule_id=rule_ids[2],
    isEnabled=False
)
print(res3)

# Test 4: Change actions to moveToFolder folders_id[0]
print("\nTest 4: Change actions (moveToFolder) for rule_id=0")
res4 = outlookMail_update_message_rule(
    rule_id=rule_ids[0],
    actions={"moveToFolder": folder_ids[0]}
)
print(res4)

# Test 5: Change actions to moveToFolder folders_id[1] + markAsRead
print("\nTest 5: Change actions (moveToFolder + markAsRead) for rule_id=1")
res5 = outlookMail_update_message_rule(
    rule_id=rule_ids[1],
    actions={
        "moveToFolder": folder_ids[1],
        "markAsRead": True
    }
)
print(res5)

# Test 6: Change conditions only
print("\nTest 6: Change conditions (subjectContains) for rule_id=2")
res6 = outlookMail_update_message_rule(
    rule_id=rule_ids[2],
    conditions={
        "subjectContains": ["update test"]
    }
)
print(res6)

# Test 7: Change exceptions only
print("\nTest 7: Change exceptions (fromAddresses) for rule_id=0")
res7 = outlookMail_update_message_rule(
    rule_id=rule_ids[0],
    exceptions={
        "fromAddresses": dummy_recipient
    }
)
print(res7)

# Test 8: Full update with everything (rule_id=1)
print("\nTest 8: Full update for rule_id=1")
res8 = outlookMail_update_message_rule(
    rule_id=rule_ids[1],
    displayName="Fully Updated Rule",
    sequence=10,
    isEnabled=True,
    actions={
        "moveToFolder": folder_ids[2],
        "markImportance": "high",
        "stopProcessingRules": True
    },
    conditions={
        "importance": "high",
        "fromAddresses": dummy_recipient
    },
    exceptions={
        "subjectContains": ["skip"]
    }
)
print(res8)

# Test 9: Clear conditions (match all)
print("\nTest 9: Clear conditions for rule_id=2")
res9 = outlookMail_update_message_rule(
    rule_id=rule_ids[2],
    conditions={}
)
print(res9)

# Test 10: Clear exceptions
print("\nTest 10: Clear exceptions for rule_id=0")
res10 = outlookMail_update_message_rule(
    rule_id=rule_ids[0],
    exceptions={}
)
print(res10)

print(outlookMail_delete_message_rule(rule_ids[1]))
print(outlookMail_delete_message_rule(rule_ids[2]))


print("\n----------------- Testing list messages extra -----------------\n")

message_ids = []

print("\nTest 4.1: List messages from main mailbox")
list_result = outlookMail_list_messages(
    top=5,
    select="subject,from,receivedDateTime"
)
print(list_result)

if 'value' in list_result:
    for msg in list_result['value'][:3]:
        message_ids.append(msg.get('id'))
else:
    print("Failed to fetch messages from main mailbox")

print("\nCollected message IDs after mailbox list:", message_ids)

print("\nTest 4.2: List messages from folder")
folder_id = folder_ids[1]
folder_list_result = outlookMail_list_messages_from_folder(
    folder_id=folder_id,
    top=5,
    select="subject,from,receivedDateTime"
)
print(folder_list_result)

if 'value' in folder_list_result:
    for msg in folder_list_result['value'][:3]:
        message_ids.append(msg.get('id'))
else:
    print("Failed to fetch messages from folder")

print("\nFinal collected message IDs:", message_ids)


print("\n----------------- Testing draft create/update extra -----------------\n")

draft_ids = []

# Folder to use (use first three folder IDs you have)
folders_to_use = folder_ids[:3]

# 1. Create draft in folder 1
print("\nTest 5.1: Create draft in first folder")
result_1 = outlookMail_create_draft_in_folder(
    folder_id=folders_to_use[0],
    subject="Draft 1 - Hello",
    body_content="<p>This is test draft 1</p>",
    to_recipients=["test1@example.com"]
)
print(result_1)
draft_id_1 = result_1.get("id")
if draft_id_1: draft_ids.append(draft_id_1)

# 2. Create draft in folder 2
print("\nTest 5.2: Create draft in second folder with CC and categories")
result_2 = outlookMail_create_draft_in_folder(
    folder_id=folders_to_use[1],
    subject="Draft 2 - CC test",
    body_content="<p>Body with CC</p>",
    to_recipients=["test2@example.com"],
    cc_recipients=["cc1@example.com", "cc2@example.com"],
    categories=["ProjectX", "FollowUp"]
)
print(result_2)
draft_id_2 = result_2.get("id")
if draft_id_2: draft_ids.append(draft_id_2)

# 3. Create draft in folder 3 with high importance
print("\nTest 5.3: Create draft in third folder with High importance")
result_3 = outlookMail_create_draft_in_folder(
    folder_id=folders_to_use[2],
    subject="Draft 3 - Important",
    body_content="<p>Important draft</p>",
    to_recipients=["test3@example.com"],
    importance="High"
)
print(result_3)
draft_id_3 = result_3.get("id")
if draft_id_3: draft_ids.append(draft_id_3)

# Print collected draft IDs
print("\nCollected draft IDs:", draft_ids)

# Only run updates if drafts were created
if draft_ids:
    # 4. Simple update: change subject
    print("\nTest 5.4: Update draft 1 - change subject")
    update_1 = outlookMail_update_draft(
        message_id=draft_ids[0],
        subject="Draft 1 - Updated Subject"
    )
    print(update_1)

    # 5. Update draft 2: change body content and add BCC
    print("\nTest 5.5: Update draft 2 - change body and add BCC")
    update_2 = outlookMail_update_draft(
        message_id=draft_ids[1],
        body_content="<p>Updated body content</p>",
        bcc_recipients=["bcc1@example.com"]
    )
    print(update_2)

    # 6. Update draft 3: mark as read and lower importance
    print("\nTest 5.6: Update draft 3 - mark as read and set importance Low")
    update_3 = outlookMail_update_draft(
        message_id=draft_ids[2],
        is_read=True,
        importance="Low"
    )
    print(update_3)

    # 7. Update draft 1: set categories
    print("\nTest 5.7: Update draft 1 - set categories")
    update_4 = outlookMail_update_draft(
        message_id=draft_ids[0],
        categories=["Urgent", "Personal"]
    )
    print(update_4)

    # 8. Update draft 2: clear CC recipients (set empty list)
    print("\nTest 5.8: Update draft 2 - clear CC recipients")
    update_5 = outlookMail_update_draft(
        message_id=draft_ids[1],
        cc_recipients=[]
    )
    print(update_5)

    # 9. Update draft 3: change subject and add reply-to
    print("\nTest 5.9: Update draft 3 - change subject and add reply-to")
    update_6 = outlookMail_update_draft(
        message_id=draft_ids[2],
        subject="Draft 3 - Final Subject",
        reply_to=["replyto@example.com"]
    )
    print(update_6)

    # 10. Update draft 1: add toRecipients and flag
    print("\nTest 5.10: Update draft 1 - add extra recipient and follow-up flag")
    followup_flag = {
        "flagStatus": "flagged",
        "startDateTime": {
            "dateTime": "2025-07-15T09:00:00",
            "timeZone": "Pacific Standard Time"
        },
        "dueDateTime": {
            "dateTime": "2025-07-16T17:00:00",
            "timeZone": "Pacific Standard Time"
        }
    }
    update_7 = outlookMail_update_draft(
        message_id=draft_ids[0],
        to_recipients=["newrecipient@example.com"],
        flag=followup_flag
    )
    print(update_7)
else:
    print("Skipping update tests: drafts could not be created.")

print("\n----------------- Testing draft delete / copy / forward extra -----------------\n")

if draft_ids:
    # 1. Delete draft 1
    print("\nTest 6.1: Delete draft 1")
    delete_result_1 = outlookMail_delete_draft(draft_ids[0])
    print(delete_result_1)

    # 2. Copy draft 2 to folder_ids[0]
    print("\nTest 6.2: Copy draft 2 to folder_ids[0]")
    copy_result_2 = outlookMail_copy_message(
        message_id=draft_ids[1],
        destination_folder_id=folder_ids[0]
    )
    print(copy_result_2)

    # 3. Copy draft 3 to folder_ids[1]
    print("\nTest 6.3: Copy draft 3 to folder_ids[1]")
    copy_result_3 = outlookMail_copy_message(
        message_id=draft_ids[2],
        destination_folder_id=folder_ids[1]
    )
    print(copy_result_3)

    # 4. Create forward draft for draft 2
    print("\nTest 6.4: Create forward draft from draft 2")
    forward_result_2 = outlookMail_create_forward_draft(
        message_id=draft_ids[1],
        comment="FYI - forwarding this draft",
        to_recipients=["forwardto@example.com"]
    )
    print(forward_result_2)

    # 5. Create forward draft for draft 3
    print("\nTest 6.5: Create forward draft from draft 3")
    forward_result_3 = outlookMail_create_forward_draft(
        message_id=draft_ids[2],
        comment="Please see below",
        to_recipients=["team@example.com"]
    )
    print(forward_result_3)

else:
    print("Skipping tests: no draft_ids available.")

print("\n----------------- Testing forward / reply / reply-all -----------------\n")

if message_ids:
    # 1. Forward first message
    print("\nTest 7.1: Forward first message")
    forward_result = outlookMail_forward_message(
        message_id=message_ids[0],
        to_recipients=["mayank.msj.singh@gmail.com"],
        comment="FYI, see below"
    )
    print(forward_result)

    # 2. Create reply draft to second message
    print("\nTest 7.2: Create reply draft to second message")
    reply_draft_result = outlookMail_create_reply_draft(
        message_id=message_ids[1],
        comment="Thanks, got it!"
    )
    print(reply_draft_result)

    # 3. Create reply-all draft to third message
    print("\nTest 7.3: Create reply-all draft to third message")
    reply_all_draft_result = outlookMail_create_reply_all_draft(
        message_id=message_ids[2],
        comment="Looping everyone in"
    )
    print(reply_all_draft_result)

else:
    print("Skipping tests: no message_ids available.")
