# Bard Launcher 

The Bard Launcher GUI is a tool designed to help automate the process of launching multiple FFXIV instances for bard performances. It includes features for managing configurations and launching shortcuts, as well as some experimental functionalities.

## Features

- **Config Directory**: Allows you to select the directory where your bard configuration files are stored.
- **Shortcut Directory**: Allows you to select the directory where your XIVLauncher shortcuts are stored.
- **Seconds Delay**: Sets the delay in seconds between launching each shortcut (minimum 10 seconds).
- **Dark Mode**: Toggle between light and dark themes for the application interface.
- **Start All**: Launches all configured shortcuts with the specified delay.
- **Start Selected**: Launches only the selected shortcuts with the specified delay.
- **Move Default Config**: Moves the default configuration file to the selected config directory.
- **Copy Config for Individual Bard**: Copies the current FFXIV configuration for the selected bard.
- **LightAmp Integration**: Automatically starts LightAmp if it is not already running before launching the bards.
- **Create Shortcuts**: Creates shortcuts for selected accounts from the `accountsList.json` file.
- **Roaming Directory**: Allows the use of a roaming directory for FFXIV configurations.
- **Rename Shortcuts**: Allows renaming of shortcuts and corresponding configuration files.
- **Grid View**: Displays bard shortcuts as a grid with icons.
- **Context Menu**: Right-click on a bard shortcut to access additional options.

## Usage

### Main Tab

1. **Start All**: Launches all shortcuts with the delay specified in the "Seconds Delay" field.
2. **Start Selected**: Launches only the selected shortcuts with the delay specified in the "Seconds Delay" field.
3. **Move Default Config**: Moves the `default.cfg` file to the FFXIV configuration directory.
4. **Status**: Displays the status and logs of operations performed.

### Settings Tab

1. **Config Directory**: Use the "Browse" button to select the directory where your configuration files are stored.
2. **Shortcut Directory**: Use the "Browse" button to select the directory where your shortcuts are stored.
3. **Seconds Delay**: Set the delay in seconds between launching each shortcut (minimum 10 seconds).
4. **Dark Mode**: Toggle between light and dark themes.
5. **Save Settings**: Saves the current settings to a configuration file.
6. **Load Settings**: Loads the settings from the configuration file.
7. **Reset Configuration**: Resets the configuration to the default settings.

### Experimental Tab

1. **Shortcut Creator**
   - **accountsList.json Path**: Use the "Browse" button to select the `accountsList.json` file.
   - **Shortcut Directory**: Use the "Browse" button to select the directory where the shortcuts will be created.
   - **Use Roaming Directory**: Toggle to enable or disable the use of a roaming directory.
   - **Roaming Directory**: Use the "Browse" button to select the roaming directory (enabled only if "Use Roaming Directory" is checked).
   - **Create Shortcuts**: Creates shortcuts for the selected accounts.
2. **Run LightAmp**
   - **Run LightAmp**: Toggle to enable or disable running LightAmp before launching bards.
   - **LightAmp Location**: Use the "Browse" button to select the `LightAmp.exe` executable.

### Readme Tab

- **Readme**: Displays this readme file. Ensure that the `Readme.txt` file is in the same directory as the executable.

### Context Menu

- **Launch**: Launch the selected bard shortcut.
- **Copy Config**: Copy the current FFXIV configuration file to the selected bard's configuration file.
- **Change Icon**: Change the icon of the selected bard shortcut.
- **Rename**: Rename the selected bard shortcut and its corresponding configuration file.

## Setup

1. Place your XIVLauncher shortcuts for each bard you want to load in the shortcuts folder.
2. If you want separate configurations for your bards, the config files need to be named the same as your XIVLauncher shortcuts and placed in the config folder.
3. If you want a Default/Main configuration file, it needs to be named `default.cfg`.

## Troubleshooting

- Ensure that all paths and directories are correctly set.
- Make sure that the FFXIV config file is in the correct location.
- Check the status log for any error messages.
- This program also assumes you have a way to open more than 2 FFXIV windows (Which I won't discuss here).

## Support

For additional support, please contact Jackt on Discord.

## Disclaimer

The script provided is offered for use "as is." The author and/or provider of this script makes no warranties, express or implied, regarding its accuracy, functionality, or suitability for any particular purpose. By using this script, you acknowledge and agree that you do so entirely at your own risk.
