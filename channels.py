# Robotic process automation using Automagica (github.com/OakwoodAI/Automagica)
# for creating new Channels for the Account Management Portal in Microsoft Teams.
#
# IMPORTANT: This code is only compatible with displays having resolution 1920x1080
# WITHOUT scaling.
#
# Created by Rohit Shaw
# Version 1.0
#
# Property of NetApp Inc

from automagica import *
import pandas as pd
import pyperclip
import operator


# Create the channel
def create_channel(team, channel):
    global browser
    browser.find_element_by_xpath(team).click()
    Wait(0.5)
    PressKey("down")
    Wait(0.1)
    PressKey("down")
    Wait(0.1)
    PressKey("enter")
    Wait(1)
    pyperclip.copy(channel)
    PressHotkey("ctrl", "v")
    Wait(0.1)
    PressKey("tab")
    Wait(0.1)
    PressKey("tab")
    Wait(0.1)
    PressKey("tab")
    Wait(0.1)
    PressKey("tab")
    Wait(0.1)
    PressKey("tab")
    Wait(0.1)
    PressKey("tab")
    Wait(0.1)
    PressKey("enter")


# Add tab Workloads
def add_workloads():
    global browser
    browser.find_element_by_xpath(
        '//*[@id="app-messages-header"]/ng-include[2]/div/div/button'
    ).click()
    Wait(3)
    ClickOnPosition(959, 475)
    Wait(0.75)
    pyperclip.copy("Workloads")
    PressHotkey("ctrl", "v")
    Wait(0.2)
    PressKey("tab")
    Wait(0.2)
    pyperclip.copy("https://tableau.netapp.com")
    PressHotkey("ctrl", "v")
    Wait(2)
    PressKey("enter")


# Add tab LinkedIn Navigator
def add_linkedin_navigator():
    global browser
    browser.find_element_by_xpath(
        '//*[@id="app-messages-header"]/ng-include[2]/div/div/button'
    ).click()
    Wait(3)
    ClickOnPosition(959, 475)
    Wait(0.75)
    pyperclip.copy("LinkedIn Navigator")
    PressHotkey("ctrl", "v")
    Wait(0.2)
    PressKey("tab")
    Wait(0.2)
    pyperclip.copy("https://www.linkedin.com/sales")
    PressHotkey("ctrl", "v")
    Wait(2)
    PressKey("enter")


# Add tab MWA
def add_mwa():
    global browser
    browser.find_element_by_xpath(
        '//*[@id="app-messages-header"]/ng-include[2]/div/div/button'
    ).click()
    Wait(3)
    ClickOnPosition(959, 475)
    Wait(0.75)
    pyperclip.copy("MWA")
    PressHotkey("ctrl", "v")
    Wait(0.2)
    PressKey("tab")
    Wait(0.2)
    pyperclip.copy(
        "https://netapp.sharepoint.com/sites/BusinessOperationsEMEA/SitePages/Must-Win-Areas.aspx"
    )
    PressHotkey("ctrl", "v")
    Wait(2)
    PressKey("enter")


# Remove tab Wiki
def remove_wiki():
    RightClickOnPosition(626, 240)
    Wait(0.3)
    PressKey("down")
    Wait(0.1)
    PressKey("enter")
    Wait(0.5)
    PressKey("enter")
    Wait(1)


# Add folder AVS Tools
def add_avs_tools():
    global browser
    browser.find_element_by_xpath('//*[@id="files"]').click()
    Wait(8)
    browser.find_element_by_xpath('//*[@id="action-create-button"]').click()
    Wait(0.2)
    browser.find_element_by_xpath(
        '//*[@id="fileActionBar0"]/div/ul[2]/li[1]/button'
    ).click()
    Wait(2)
    pyperclip.copy("AVS Tools")
    PressHotkey("ctrl", "v")
    Wait(0.1)
    PressKey("enter")


# Make folder AVS Tools a tab
def avs_tools_tab():
    global browser
    RightClickOnPosition(602, 420)
    Wait(0.3)
    PressKey("up")
    Wait(0.1)
    PressKey("enter")
    Wait(2)
    browser.find_element_by_xpath(
        '//*[@id="files-link-dialog"]/div/div/div[2]/div/ul/li[2]/button'
    ).click()
    Wait(0.2)
    PressHotkey("ctrl", "c")
    link = pyperclip.paste()
    Wait(0.1)

    browser.find_element_by_xpath(
        '//*[@id="app-messages-header"]/ng-include[2]/div/div[1]/button'
    ).click()
    Wait(3)
    ClickOnPosition(959, 475)
    Wait(0.75)
    pyperclip.copy("AVS Tools")
    PressHotkey("ctrl", "v")
    Wait(0.2)
    PressKey("tab")
    Wait(0.2)
    pyperclip.copy(link)
    PressHotkey("ctrl", "v")
    Wait(2)
    PressKey("enter")


# Copy Account Plan Template
def copy_template():
    global browser
    browser.find_element_by_xpath('//*[@id="files"]').click()
    Wait(4)
    browser.find_element_by_xpath('//*[@id="files-upload-button"]').click()
    Wait(3)
    if country == "UK&I":
        pyperclip.copy(
            "C:\\Users\\rshaw\\OneDrive - NetApp Inc\\AMP Files\\UKI"
        )  # Copy files and change directory to yours
    else:
        pyperclip.copy(
            "C:\\Users\\rshaw\\OneDrive - NetApp Inc\\AMP Files\\Others"
        )  # Copy files and change directory to yours

    for i in range(0, 5):
        PressKey("tab")
        Wait(0.1)

    PressKey("enter")
    Wait(0.1)
    PressHotkey("ctrl", "v")
    Wait(0.1)
    PressKey("enter")
    Wait(0.3)

    for i in range(0, 4):
        PressKey("tab")
        Wait(0.1)

    PressKey("down")
    Wait(0.1)
    PressKey("enter")


# Make Account Plan Template a tab
def template_tab():
    RightClickOnPosition(633, 473)
    Wait(0.3)
    PressKey("up")
    Wait(0.1)
    PressKey("enter")


# Copy AVS files
def copy_avs_files():
    global browser
    browser.find_element_by_xpath('//*[@id="files"]').click()
    Wait(4)
    ClickOnPosition(602, 420)
    Wait(1)
    browser.find_element_by_xpath('//*[@id="files-upload-button"]').click()
    Wait(2)

    for i in range(0, 5):
        PressKey("tab")
        Wait(0.1)

    PressKey("enter")
    Wait(0.1)
    if country == "UK&I":
        pyperclip.copy(
            "C:\\Users\\rshaw\\OneDrive - NetApp Inc\\AMP Files\\UKI"
        )  # Copy files and change directory to yours
    else:
        pyperclip.copy(
            "C:\\Users\\rshaw\\OneDrive - NetApp Inc\\AMP Files\\Others"
        )  # Copy files and change directory to yours
    PressHotkey("ctrl", "v")
    Wait(0.1)
    PressKey("enter")
    Wait(0.2)

    for i in range(0, 4):
        PressKey("tab")
        Wait(0.1)

    PressKey("down")
    Wait(0.1)
    PressKey("enter")
    Wait(0.2)
    PressHotkey("ctrl", "a")
    Wait(0.1)
    PressKey("enter")


# Sequence of actions to perform
def actions():
    Wait(4)
    add_workloads()
    Wait(4)
    add_linkedin_navigator()
    Wait(4)
    add_mwa()
    Wait(8)
    add_avs_tools()
    Wait(4)
    avs_tools_tab()
    Wait(4)
    copy_template()
    Wait(15)
    template_tab()
    Wait(1)
    copy_avs_files()
    Wait(4)
    remove_wiki()


# Check if adding a channel is possible or not - Teams with single channel
def channel_counter_1(country, team):
    global channel_flag
    if 190 <= team < 200:
        print("Low on channel spots. Create a new team and reconfigure soon.")
        print(country + ": " + str(team))
        channel_flag = 0

    if team == 200:
        print(
            "No channel spots available in one or more teams. Create a new team for "
            + country
            + " and reconfigure."
        )
        print(country + ": " + str(team))
        channel_flag = 1


# Check if adding a channel is possible or not - Teams with two channels
def channel_counter_2(country, team1, team2):
    global channel_flag
    if 190 <= team1 < 200 or 190 <= team2 < 200:
        print("Low on channel spots. Create a new team and reconfigure soon.")
        print(country + "1: " + str(team1))
        print(country + "2: " + str(team2))
        channel_flag = 0

    if team1 == 200 or team2 == 200:
        print(
            "No channel spots available in one or more teams. Create a new team for "
            + country
            + " and reconfigure."
        )
        print(country + "1: " + str(team1))
        print(country + "2: " + str(team2))
        channel_flag = 1


# Check if adding a channel is possible or not - Teams with three channels
def channel_counter_3(country, team1, team2, team3):
    global channel_flag
    if 190 <= team1 < 200 or 190 <= team2 < 200 or 190 <= team3 < 200:
        print("Low on channel spots. Create a new team and reconfigure soon.")
        print(country + "1: " + str(team1))
        print(country + "2: " + str(team2))
        print(country + "3: " + str(team3))
        channel_flag = 0

    if team1 == 200 or team2 == 200 or team3 == 200:
        print(
            "No channel spots available in one or more teams. Create a new team for "
            + country
            + " and reconfigure."
        )
        print(country + "1: " + str(team1))
        print(country + "2: " + str(team2))
        print(country + "3: " + str(team3))
        channel_flag = 1


# Select the best team to add the new channel to, ensuring a balance between Countries with multiple teams
def team_selector(country):
    global team_selector_out
    a_dict = {}
    df = pd.read_csv("AMP Teams and Channels.csv", index_col=0,)

    if country == "Netherlands":
        NL1 = df.loc["Netherlands1", "Channels"]
        NL2 = df.loc["Netherlands2", "Channels"]
        channel_counter_2(country, NL1, NL2)
        if channel_flag == 0:
            if NL1 <= NL2:
                NL1 = NL1 + 1
                df.at["Netherlands1", "Channels"] = NL1
                team_selector_out = netherlands1
            else:
                NL2 = NL2 + 1
                df.at["Netherlands2", "Channels"] = NL2
                team_selector_out = netherlands2

    elif country == "France":
        France1 = df.loc["France1", "Channels"]
        France2 = df.loc["France2", "Channels"]
        channel_counter_2(country, France1, France2)
        if channel_flag == 0:
            if France1 <= France2:
                France1 = France1 + 1
                df.at["France1", "Channels"] = France1
                team_selector_out = france1
            else:
                France2 = France2 + 1
                df.at["France2", "Channels"] = France2
                team_selector_out = france2

    elif country == "Germany":
        Germany1 = df.loc["Germany1", "Channels"]
        Germany2 = df.loc["Germany2", "Channels"]
        Germany3 = df.loc["Germany3", "Channels"]
        a_dict["Germany1"] = Germany1
        a_dict["Germany2"] = Germany2
        a_dict["Germany3"] = Germany3
        min_val = min(a_dict.items(), key=operator.itemgetter(1))[0]
        min_channel_count = min(a_dict.items(), key=operator.itemgetter(1))[1]

        channel_counter_3(country, Germany1, Germany2, Germany3)
        if channel_flag == 0:
            if min_val == "Germany1":
                min_channel_count = min_channel_count + 1
                df.at["Germany1", "Channels"] = min_channel_count
                team_selector_out = germany1
            elif min_val == "Germany2":
                min_channel_count = min_channel_count + 1
                df.at["Germany2", "Channels"] = min_channel_count
                team_selector_out = germany2
            elif min_val == "Germany3":
                min_channel_count = min_channel_count + 1
                df.at["Germany3", "Channels"] = min_channel_count
                team_selector_out = germany3
            else:
                print("*A wild error appears in Germany*")
                PressHotkey("ctrl", "c")
                PressHotkey("ctrl", "c")

    elif country == "UK&I":
        UKI1 = df.loc["UKI1", "Channels"]
        UKI2 = df.loc["UKI2", "Channels"]
        UKI3 = df.loc["UKI3", "Channels"]
        a_dict["UKI1"] = UKI1
        a_dict["UKI2"] = UKI2
        a_dict["UKI3"] = UKI3
        min_val = min(a_dict.items(), key=operator.itemgetter(1))[0]
        min_channel_count = min(a_dict.items(), key=operator.itemgetter(1))[1]

        channel_counter_3(country, UKI1, UKI2, UKI3)
        if channel_flag == 0:
            if min_val == "UKI1":
                min_channel_count = min_channel_count + 1
                df.at["UKI1", "Channels"] = min_channel_count
                team_selector_out = uki1
            elif min_val == "UKI2":
                min_channel_count = min_channel_count + 1
                df.at["UKI2", "Channels"] = min_channel_count
                team_selector_out = uki2
            elif min_val == "UKI3":
                min_channel_count = min_channel_count + 1
                df.at["UKI3", "Channels"] = min_channel_count
                team_selector_out = uki3
            else:
                print("*A wild error appears in UK&I*")
                PressHotkey("ctrl", "c")
                PressHotkey("ctrl", "c")

    elif country == "Russia":
        Russia1 = df.loc["Russia1", "Channels"]
        Russia2 = df.loc["Russia2", "Channels"]
        channel_counter_2(country, Russia1, Russia2)
        if channel_flag == 0:
            if Russia1 <= Russia2:
                Russia1 = Russia1 + 1
                df.at["Russia1", "Channels"] = Russia1
                team_selector_out = russia1
            else:
                Russia2 = Russia2 + 1
                df.at["Russia2", "Channels"] = Russia2
                team_selector_out = russia2

    elif country == "MiddleEast":
        ME1 = df.loc["MiddleEast1", "Channels"]
        ME2 = df.loc["MiddleEast2", "Channels"]
        channel_counter_2(country, ME1, ME2)
        if channel_flag == 0:
            if ME1 <= ME2:
                ME1 = ME1 + 1
                df.at["MiddleEast1", "Channels"] = ME1
                team_selector_out = middleeast1
            else:
                ME2 = ME2 + 1
                df.at["MiddleEast2", "Channels"] = ME2
                team_selector_out = middleeast2

    elif country == "Spain":
        channels = df.loc["Spain", "Channels"]
        channel_counter_1(country, channels)
        if channel_flag == 0:
            channels = channels + 1
            df.at["Spain", "Channels"] = channels
            team_selector_out = spain

    elif country == "Israel":
        channels = df.loc["Israel", "Channels"]
        channel_counter_1(country, channels)
        if channel_flag == 0:
            channels = channels + 1
            df.at["Israel", "Channels"] = channels
            team_selector_out = israel

    elif country == "Switzerland":
        channels = df.loc["Switzerland", "Channels"]
        channel_counter_1(country, channels)
        if channel_flag == 0:
            channels = channels + 1
            df.at["Switzerland", "Channels"] = channels
            team_selector_out = switzerland

    elif country == "Sweden":
        channels = df.loc["Sweden", "Channels"]
        channel_counter_1(country, channels)
        if channel_flag == 0:
            channels = channels + 1
            df.at["Sweden", "Channels"] = channels
            team_selector_out = sweden

    elif country == "Italy":
        channels = df.loc["Italy", "Channels"]
        channel_counter_1(country, channels)
        if channel_flag == 0:
            channels = channels + 1
            df.at["Italy", "Channels"] = channels
            team_selector_out = italy

    elif country == "Austria":
        channels = df.loc["Austria", "Channels"]
        channel_counter_1(country, channels)
        if channel_flag == 0:
            channels = channels + 1
            df.at["Austria", "Channels"] = channels
            team_selector_out = austria

    else:
        print("*A wild error appears at team_selector*")
        PressHotkey("ctrl", "c")
        PressHotkey("ctrl", "c")

    df.to_csv("AMP Teams and Channels.csv")
    a_dict.clear()


# Select the channel based on entered country
def channel_selector():
    global country

    if country == "test":
        pass

    elif country in ["Netherlands", "netherlands", "NETHERLANDS"]:
        country = "Netherlands"
        team_selector(country)

    elif country in ["France", "france", "FRANCE"]:
        country = "France"
        team_selector(country)

    elif country in ["Spain", "spain", "SPAIN"]:
        country = "Spain"
        team_selector(country)

    elif country in ["Germany", "germany", "GERMANY"]:
        country = "Germany"
        team_selector(country)

    elif country in ["Israel", "israel", "ISRAEL"]:
        country = "Israel"
        team_selector(country)

    elif country in ["UK", "UK&I", "UKI", "uk", "uk&i", "uki"]:
        country = "UK&I"
        team_selector(country)

    elif country in ["Switzerland", "switzerland", "SWITZERLAND"]:
        country = "Switzerland"
        team_selector(country)

    elif country in ["Sweden", "sweden", "SWEDEN"]:
        country = "Sweden"
        team_selector(country)

    elif country in ["Italy", "italy", "ITALY"]:
        country = "Italy"
        team_selector(country)

    elif country in ["Russia", "russia", "RUSSIA"]:
        country = "Russia"
        team_selector(country)

    elif country in [
        "Middle East",
        "MiddleEast",
        "middle east",
        "middleeast",
        "MIDDLEEAST",
        "MIDDLE EAST",
    ]:
        country = "MiddleEast"
        team_selector(country)

    elif country in ["Austria", "austria", "AUSTRIA"]:
        country = "Austria"
        team_selector(country)

    else:
        print("*A wild error appears at operations*")
        PressHotkey("ctrl", "c")
        PressHotkey("ctrl", "c")


# Operations to perform on the browser after login
def operations():
    global country
    global channel_name
    global team_selector_out

    if country == "test":
        create_channel(test, channel_name)
        actions()

    elif country in [
        "Netherlands",
        "France",
        "Spain",
        "Germany",
        "Israel",
        "UK&I",
        "Switzerland",
        "Sweden",
        "Italy",
        "Russia",
        "MiddleEast",
        "Austria",
    ]:
        create_channel(team_selector_out, channel_name)
        actions()

    else:
        print("*A wild error appears at operations*")
        PressHotkey("ctrl", "c")
        PressHotkey("ctrl", "c")

    country = ""


# Check if entered country is valid or not
def country_check():
    global country
    global action_flag

    if country in [
        "Netherlands",
        "netherlands",
        "NETHERLANDS",
        "France",
        "france",
        "FRANCE",
        "Spain",
        "spain",
        "SPAIN",
        "Germany",
        "germany",
        "GERMANY",
        "Israel",
        "israel",
        "ISRAEL",
        "UK",
        "UK&I",
        "UKI",
        "uk",
        "uk&i",
        "uki",
        "Switzerland",
        "switzerland",
        "SWITZERLAND",
        "Sweden",
        "sweden",
        "SWEDEN",
        "Italy",
        "italy",
        "ITALY",
        "Russia",
        "russia",
        "RUSSIA",
        "Middle East",
        "MiddleEast",
        "middle east",
        "middleeast",
        "MIDDLEEAST",
        "MIDDLE EAST",
        "Austria",
        "austria",
        "AUSTRIA",
        "test",
    ]:
        action_flag = 1

    else:
        print("Country not found. Try again.")
        channel_details()


# Input - Details of channel to be created
def channel_details():
    global channel_name
    global country
    global action_flag
    print("\n----------------------------------------------")
    channel_name = input("Channel Name: ")
    country = input("Country: ")
    country_check()
    if action_flag == 1:
        channel_selector()
    else:
        print("*A wild error appears at channel_details*")
        PressHotkey("ctrl", "c")
        PressHotkey("ctrl", "c")


def main():
    global browser

    # Start
    user = input("Email: ")
    channel_details()

    if channel_flag == 0:
        print("\nLaunch the Authenticator app on your phone!")
        print("----------------------------------------------")

        # Open Chrome
        browser = ChromeBrowser()
        DoubleClickOnPosition(458, 30)
        browser.get("https://teams.microsoft.com")

        # Login
        Wait(0.5)
        username = browser.find_element_by_name("loginfmt")
        username.send_keys(user)
        PressKey("enter")
        Wait(0.1)
        PressKey("enter")
        Wait(20)

        PressKey("tab")
        PressKey("enter")

        Wait(30)

        # Click on Teams
        ClickOnPosition(35, 299)

        operations()


global channel_name
global country
global browser
global action_flag
global team_selector_out
global channel_flag

channel_name = ""
country = ""
action_flag = 0
channel_flag = 0

# Teams and their IDs
test = '//*[@id="team-19:1659abc31cda45ff915bf8183a1eda76@thread.skype"]/a/button'
netherlands2 = (
    '//*[@id="team-19:0185db89865c472ea8c8751baa7c06f6@thread.skype"]/a/button'
)
france1 = '//*[@id="team-19:0b340e695b1243848208380e11fd87f0@thread.skype"]/a/button'
spain = '//*[@id="team-19:0e5adfa1c5be43ca87cf02495436d2f7@thread.skype"]/a/button'
germany3 = '//*[@id="team-19:3e50b78f2e5f4d68a38397e28a75a0d0@thread.skype"]/a/button'
israel = '//*[@id="team-19:4491c08102a54762ad4fa9cef9a84303@thread.skype"]/a/button'
uki2 = '//*[@id="team-19:5321d5b009c44294a4050b207355e883@thread.skype"]/a/button'
switzerland = (
    '//*[@id="team-19:6463379af9214e9e978bb170a1b478c8@thread.skype"]/a/button'
)
sweden = '//*[@id="team-19:801bd7c49fc04122b18173b81e33d6b4@thread.skype"]/a/button'
uki1 = '//*[@id="team-19:87c6f6c530224b7ea33e27eb3e81bfb7@thread.skype"]/a/button'
germany1 = '//*[@id="team-19:891ebab877e7419e80b76e3b871772cb@thread.skype"]/a/button'
italy = '//*[@id="team-19:9538afc06374461d86a4a6c6f99b6211@thread.skype"]/a/button'
germany2 = '//*[@id="team-19:99c4af8b693a4c0fbc5aa74fd8801ecb@thread.skype"]/a/button'
uki3 = '//*[@id="team-19:b5509d09b2c941279c0c72923bbf7aec@thread.skype"]/a/button'
russia2 = '//*[@id="team-19:db89f89ba185432daab3b77cd093e3e0@thread.skype"]/a/button'
france2 = '//*[@id="team-19:e258dfc4a517453f9dcc6516a95c662a@thread.skype"]/a/button'
middleeast2 = (
    '//*[@id="team-19:eae581eaae1148a4b0301aaf2184322a@thread.skype"]/a/button'
)
austria = '//*[@id="team-19:f59e0943dd734921a837d47488b2f534@thread.skype"]/a/button'
russia1 = '//*[@id="team-19:f63c765e89c64194b6ffb2ce0b61de08@thread.skype"]/a/button'
netherlands1 = (
    '//*[@id="team-19:fb8fb226562041d794f0372eeee0bac3@thread.skype"]/a/button'
)
middleeast1 = (
    '//*[@id="team-19:fc66301e28354a798d9487c1e9b46a50@thread.skype"]/a/button'
)
