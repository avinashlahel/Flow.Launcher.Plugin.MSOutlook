import os
import sys
from datetime import datetime, timedelta

parent_folder_path = os.path.abspath(os.path.dirname(__file__))
sys.path.append(os.path.join(parent_folder_path, 'lib'))
sys.path.append(os.path.join(parent_folder_path, 'plugin'))
sys.path.append(parent_folder_path)

# If pywin32 is vendored, add its DLL folder to the search path before importing win32com
pywin32_system32 = os.path.join(parent_folder_path, 'lib', 'pywin32_system32')
if os.path.isdir(pywin32_system32):
    try:
        # Python 3.8+: preferred way
        os.add_dll_directory(pywin32_system32)
    except Exception:
        # Fallback: prepend to PATH
        os.environ["PATH"] = pywin32_system32 + os.pathsep + os.environ.get("PATH", "")


from flowlauncher import FlowLauncher
import webbrowser

import win32com.client


class MSOutlook(FlowLauncher):
    def query(self, query):
        return self.get_today_meetings()

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
                "IcoPath": "Images/cal.png",
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
