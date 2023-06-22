import pandas as pd
import imaplib
import email


class Mail:
    def __init__(self, username, password, is_spam_trap):
        self.username = username
        self.password = password
        self.is_spam_trap = is_spam_trap


def count_email_addresses(lst):
    count_dict = {}
    for elem in lst:
        if elem in count_dict:
            count_dict[elem] += 1
        else:
            count_dict[elem] = 1
    return list(count_dict.items())


statistics_list = []
spam_trap_mail_list = []
domain_mail_list = []


spam_trap_mail_list.append(Mail("email_address@here.com", "emailpassword", True))

domain_mail_list.append(Mail("email_address@here.com", "emailpassword", False))


# find all spam emails on every spam trap email on list
for spam_trap_mail in spam_trap_mail_list:
    print(f"CHECKING: {spam_trap_mail.username}")

    # Connect to outlooks's IMAP server
    imap_server = imaplib.IMAP4_SSL("outlook.office365.com")

    # Log in to the server
    imap_server.login(spam_trap_mail.username, spam_trap_mail.password)

    # Get all mailboxes
    final_mailbox_list = []
    status, mailbox_list = imap_server.list()

    if status == "OK":
        # Print only the mailbox names
        for mailbox in mailbox_list:
            # Extract the mailbox name from the response
            final_mailbox_list.append(mailbox.decode().split(' "/" ')[1].strip(' "'))

    # Get all emails from every mailbox
    for mailbox in final_mailbox_list:
        imap_server.select(mailbox)

        # Search for all email messages
        status, data = imap_server.search(None, "ALL")

        if status == "OK":
            # Get the list of email IDs
            email_ids = data[0].split()

            for email_id in email_ids:
                # Fetch the email data for each ID
                status, msg_data = imap_server.fetch(email_id, "(RFC822)")

                if status == "OK":
                    # Parse the email data using the email module
                    email_msg = email.message_from_bytes(msg_data[0][1])

                    # Get the sender and subject
                    sender = email.utils.parseaddr(email_msg["From"])[1]
                    subject = email_msg["Subject"]

                    statistics_list.append(sender)

                    # print("From:", sender)
                    # print("Subject:", subject)
                    # print()

    print()

    # Log out and close the connection
    imap_server.logout()


df = pd.DataFrame(
    count_email_addresses(statistics_list), columns=["Email address", "Count"]
)
df.sort_values(by="Count", ascending=False, inplace=True)
df.to_excel("Spam_trap_statistics.xlsx", index=False)


# add every email address from spam traps to black list of domain emails
for domain_mail in domain_mail_list:
    # Connect to outlooks's IMAP server
    imap_server = imaplib.IMAP4_SSL("outlook.office365.com")

    # Log in to the email account
    imap_server.login(domain_mail.username, domain_mail.password)

    # Select the mailbox to work with (e.g., 'INBOX')
    imap_server.select("INBOX")

    for email_to_block in set(statistics_list):
        search_criteria = f'(FROM "{email_to_block}")'
        _, email_ids = imap_server.search(None, search_criteria)

        # Move each email to the Junk mailbox
        for email_id in email_ids[0].split():
            # Move the email by copying it to the Junk mailbox and then marking it as deleted
            imap_server.copy(email_id, "Junk")
            imap_server.store(email_id, "+FLAGS", "\\Deleted")
            print(f"Email with ID {email_id} has been moved to the Junk mailbox.")

    # Log out and close the connection
    imap_server.logout()
