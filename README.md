# NetApp Account Managament Portal - Channel Creator

Robotic process automation using [Automagica](github.com/OakwoodAI/Automagica) for creating new Channels for the Account Management Portal in Microsoft Teams.

## Requirements
This application requires [Python 3.7](https://www.python.org) or above and list of libraries and modules. To install the libraries, run the following command.
<pre>   pip install -r requirements.txt </pre>

Create a local copy of files to be copied to the channel.
* Create a directory for UK&I, download a copy of the template file and AVS Tools from any of the UK&I Teams.
* Create a directory for Others, download a copy of the template file and AVS Tools from a Team other than UK&I.

Update the path to directories in _copy_template()_ and _copy_avs_files()_ in _channels.py_ with the ones you just created.

**IMPORTANT: This program requires a display with resolution 1920 x 1080 without scaling. The display scaling must be set to 100%.**

## Usage
1. Run main.py
<pre>   python main.py    </pre>
2. Enter the name of the channel to be created
3. Enter the name of the country where the channel belongs
4. Enter your (short) NetApp email address
5. Launch the Authenticator app on your phone and approve login
6. Watch the channel being created!

**Do not use your computer until the channel creation process has completed.**


***Property of NetApp***
