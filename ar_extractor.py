import win32com.client
import datetime
import re
import pandas as pd
from collections import defaultdict

def extract_ar_metrics():
    # Connect to the local Outlook application
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Specify the exact name as it appears in the Outlook sidebar
    target_account = "ar@forwardartists.com" 
    inbox = None

    # Bind strictly to the target mailbox MAPI store
    for folder in outlook.Folders:
        if target_account.lower() == folder.Name.lower():
            try:
                inbox = folder.Folders("Inbox")
            except Exception as e:
                print(f"Found the account, but could not access the Inbox: {e}")
            break

    # Hard stop if the specific AR inbox isn't found
    if not inbox:
        raise ValueError(f"CRITICAL STOP: Inbox for '{target_account}' not found. Check the exact display name in your Outlook sidebar.")

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True) # Sort descending to get latest emails first

    # Restrict to the last 7 days
    week_ago = datetime.datetime.now() - datetime.timedelta(days=7)
    week_ago_str = week_ago.strftime('%m/%d/%Y %H:%M %p')
    recent_messages = messages.Restrict(f"[ReceivedTime] >= '{week_ago_str}'")

    # Aggregate thread data by Invoice Number (looks for 4-6 digits, optional "INV-" prefix)
    invoice_data = defaultdict(lambda: {'count': 0, 'senders': set(), 'last_update': None, 'summary': ''})
    invoice_regex = r'(?:INV[-\s]?)?(\d{4,6})'

    for msg in recent_messages:
        try:
            subject = msg.Subject
            match = re.search(invoice_regex, subject, re.IGNORECASE)
            
            if match:
                inv_num = match.group(1)
                invoice_data[inv_num]['count'] += 1
                invoice_data[inv_num]['senders'].add(msg.SenderName)
                
                # Capture the most recent body text for the summary snippet
                if not invoice_data[inv_num]['summary']:
                    invoice_data[inv_num]['summary'] = msg.Body[:200].strip() + "..."
                    invoice_data[inv_num]['last_update'] = msg.ReceivedTime.strftime('%Y-%m-%d %H:%M')
        except Exception:
            # Skips items that aren't standard MailItems (like calendar invites)
            continue

    # Format the output into a DataFrame
    records = []
    for inv, data in invoice_data.items():
        records.append({
            'Invoice Number': inv,
            'Total Thread Emails': data['count'],
            'Participants': ", ".join(data['senders']),
            'Last Update': data['last_update'],
            'Latest Interaction Snippet': data['summary'].replace('\r', '').replace('\n', ' ')
        })

    df = pd.DataFrame(records)
    
    # Export to CSV
    output_filename = f"AR_Weekly_Metrics_{datetime.datetime.now().strftime('%Y%m%d')}.csv"
    df.to_csv(output_filename, index=False)
    print(f"Extraction complete. Found recent activity for {len(df)} invoices. Saved to {output_filename}")

if __name__ == "__main__":
    extract_ar_metrics()