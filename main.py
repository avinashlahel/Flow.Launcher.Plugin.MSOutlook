import os
import sys
from datetime import datetime, timedelta

import win32com.client

parent_folder_path = os.path.abspath(os.path.dirname(__file__))
sys.path.append(parent_folder_path)
sys.path.append(os.path.join(parent_folder_path, 'lib'))
sys.path.append(os.path.join(parent_folder_path, 'plugin'))

from flowlauncher import FlowLauncher
import webbrowser


class MSOutlook(FlowLauncher):
    def query(self, query):
        if query.strip().lower() == "mt":
            return self.get_today_meetings()

        return []

    def get_today_meetings(self):
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(9)
        items = calendar.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        # Define the start and end of today
        today = datetime.today()
        day_start = datetime(today.year, today.month, today.day, 0, 0, 0)
        end = day_start + timedelta(days=1)
        meetings_from = datetime.now() - timedelta(minutes=15)

        restriction = "[Start] >= '{}' AND [Start] < '{}'".format(
            meetings_from.strftime("%m/%d/%Y %H:%M %p"),
            end.strftime("%m/%d/%Y %H:%M %p")
        )

        restricted_items = items.Restrict(restriction)

        results = []
        for item in restricted_items:
            now = datetime.now()
            meeting_start = item.Start.replace(tzinfo=None)
            time_diff = meeting_start - now

            if time_diff.total_seconds() < 0:
                time_status = "Meeting in progress" if now < item.End.replace(tzinfo=None) else "Meeting ended"
            else:
                hours = int(time_diff.total_seconds() // 3600)
                minutes = int((time_diff.total_seconds() % 3600) // 60)
                time_status = f"Starts in {hours}h {minutes}m" if hours > 0 else f"Starts in {minutes}m"

            results.append({
                "Title": item.Subject,
                "SubTitle": f"{time_status} ({item.Start.strftime('%H:%M')} - {item.End.strftime('%H:%M')})",
                "IcoPath": "Images/ol.png",
                "JsonRPCAction": {
                    "method": "open_selected_meeting",
                    "parameters": [item.EntryID],
                }
            })

        return results

    def open_selected_meeting(self, entry_id):
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        item = namespace.GetItemFromID(entry_id)
        item.Display()

    def open_url(self, url):
        webbrowser.open(url)


if __name__ == "__main__":
    MSOutlook()
