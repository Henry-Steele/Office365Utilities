import requests
import pandas as pd
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Helper Functions
def getAccessToken(tenantId, clientId, clientSecret, scope):
    tokenUrl = f'https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token'
    data = {
        'grant_type': 'client_credentials',
        'client_id': clientId,
        'client_secret': clientSecret,
        'scope': scope
    }
    try:
        response = requests.post(tokenUrl, data=data)
        response.raise_for_status()
        return response.json().get('access_token')
    except requests.exceptions.RequestException as e:
        raise Exception(f"Error obtaining access token: {e}")

def fetchAllUsers(accessToken):
    graphUrl = 'https://graph.microsoft.com/v1.0/users?$select=userPrincipalName,accountEnabled'
    headers = {
        'Authorization': f'Bearer {accessToken}'
    }
    try:
        users = []
        response = requests.get(graphUrl, headers=headers)
        response.raise_for_status()
        data = response.json()
        users.extend(data['value'])
        while '@odata.nextLink' in data:
            response = requests.get(data['@odata.nextLink'], headers=headers)
            response.raise_for_status()
            data = response.json()
            users.extend(data['value'])
        return users
    except requests.exceptions.RequestException as e:
        raise Exception(f"Error fetching users: {e}")

def fetchSignInLogs(accessToken):
    graphUrl = 'https://graph.microsoft.com/v1.0/auditLogs/signIns'
    headers = {
        'Authorization': f'Bearer {accessToken}'
    }
    try:
        response = requests.get(graphUrl, headers=headers)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        raise Exception(f"Error fetching sign-in logs: {e}")

def parseSignInLogs(signInLogs):
    signInData = []
    for log in signInLogs.get('value', []):
        signInData.append({
            'userPrincipalName': log.get('userPrincipalName'),
            'lastSignInDate': pd.to_datetime(log.get('createdDateTime')).date(),
            'lastSignInTime': pd.to_datetime(log.get('createdDateTime')).time()
        })
    
    df = pd.DataFrame(signInData)
    df = df.sort_values(by=['userPrincipalName', 'lastSignInDate', 'lastSignInTime']).drop_duplicates(subset=['userPrincipalName'], keep='last')
    
    return df

def normalizeFilePath(filePath):
    return os.path.normpath(filePath)

def writeDataFrameToCSV(df, filePath):
    try:
        df.to_csv(filePath, index=False)
    except IOError as e:
        raise Exception(f"Error writing to file {filePath}: {e}")

# Function to display DataFrame in a new Tkinter window
def displayDataFrame(df):
    window = tk.Toplevel(root)
    window.title("User Sign-In Logs")
    window.geometry("900x600")  # Increase window size
    frame = ttk.Frame(window, padding="10 10 10 10")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    
    tree = ttk.Treeview(frame, columns=list(df.columns), show='headings', height=20)  # Increase the height of the Treeview
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=300, anchor=tk.W)  # Set column width to 300 and left align
        
    for row in df.itertuples(index=False):
        tree.insert("", tk.END, values=row)
    
    tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# GUI Function
def runApp():
    try:
        tenantId = tenantIdEntry.get()
        clientId = clientIdEntry.get()
        clientSecret = clientSecretEntry.get()
        outputOption = outputOptionVar.get()
        outputFile = filePathEntry.get() if outputOption in ['csv file', 'both'] else None
        showDisabled = showDisabledVar.get()

        accessToken = getAccessToken(tenantId, clientId, clientSecret, scope='https://graph.microsoft.com/.default')
        
        users = fetchAllUsers(accessToken)
        user_df = pd.DataFrame(users)
        user_df['userPrincipalName'] = user_df['userPrincipalName'].str.lower()

        signInLogs = fetchSignInLogs(accessToken)
        signIn_df = parseSignInLogs(signInLogs)

        # Merge user data with sign-in logs
        result_df = pd.merge(user_df[['userPrincipalName', 'accountEnabled']], signIn_df, on='userPrincipalName', how='left')
        result_df['lastSignInDate'] = result_df['lastSignInDate'].fillna('-')
        result_df['lastSignInTime'] = result_df['lastSignInTime'].fillna('-')
        result_df['accountEnabled'] = result_df['accountEnabled'].apply(lambda x: 'Active' if x else 'Inactive')

        # Filter out disabled users if required
        if not showDisabled:
            result_df = result_df[result_df['accountEnabled'] == 'Active']

        if outputOption in ['csv file', 'both'] and outputFile:
            outputFilePath = normalizeFilePath(outputFile)
            writeDataFrameToCSV(result_df, outputFilePath)
            messagebox.showinfo("Success", f"Data exported to {outputFilePath}")
        
        if outputOption in ['table', 'both']:
            displayDataFrame(result_df)
        
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Function to update file path entry visibility
def updateFilePathEntry(*args):
    if outputOptionVar.get() in ['csv file', 'both']:
        filePathLabel.grid()
        filePathEntry.grid()
        browseButton.grid()
    else:
        filePathLabel.grid_remove()
        filePathEntry.grid_remove()
        browseButton.grid_remove()

# Function to browse file path
def browseFilePath():
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
    filePathEntry.delete(0, tk.END)
    filePathEntry.insert(0, file_path)

# GUI Setup
root = tk.Tk()
root.title("Azure AD Sign-In Logs Fetcher")

mainframe = ttk.Frame(root, padding="10 10 10 10")
mainframe.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Tenant ID
ttk.Label(mainframe, text="Tenant ID:").grid(row=0, column=0, sticky=tk.W)
tenantIdEntry = ttk.Entry(mainframe, width=30)
tenantIdEntry.grid(row=0, column=1, sticky=(tk.W, tk.E))

# Client ID
ttk.Label(mainframe, text="Client ID:").grid(row=1, column=0, sticky=tk.W)
clientIdEntry = ttk.Entry(mainframe, width=30)
clientIdEntry.grid(row=1, column=1, sticky=(tk.W, tk.E))

# Client Secret
ttk.Label(mainframe, text="Client Secret:").grid(row=2, column=0, sticky=tk.W)
clientSecretEntry = ttk.Entry(mainframe, width=30, show="*")
clientSecretEntry.grid(row=2, column=1, sticky=(tk.W, tk.E))

# Output Option
ttk.Label(mainframe, text="Output Option:").grid(row=3, column=0, sticky=tk.W)
outputOptionVar = tk.StringVar()
outputOptionMenu = ttk.Combobox(mainframe, textvariable=outputOptionVar)
outputOptionMenu['values'] = ('table', 'csv file', 'both')
outputOptionMenu.grid(row=3, column=1, sticky=(tk.W, tk.E))
outputOptionMenu.bind('<<ComboboxSelected>>', updateFilePathEntry)

# File Path
filePathLabel = ttk.Label(mainframe, text="File Path:")
filePathLabel.grid(row=4, column=0, sticky=tk.W)
filePathEntry = ttk.Entry(mainframe, width=30)
filePathEntry.grid(row=4, column=1, sticky=(tk.W, tk.E))

# Browse Button
browseButton = ttk.Button(mainframe, text="Browse", command=browseFilePath)
browseButton.grid(row=4, column=2, sticky=tk.W)

# Show Disabled Users Option
showDisabledVar = tk.BooleanVar(value=True)
showDisabledCheck = ttk.Checkbutton(mainframe, text="Show Disabled Users", variable=showDisabledVar)
showDisabledCheck.grid(row=5, column=0, columnspan=2, sticky=tk.W)

# Run Button
runButton = ttk.Button(mainframe, text="Run", command=runApp)
runButton.grid(row=6, column=1, sticky=tk.E)

# Initial State
updateFilePathEntry()

# Start the GUI
root.mainloop()