#!/bin/sh

# Laeeq Humam Twitter Handle @laeeqhumam
# Created for XWS 20160820

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# Version Control 20160929 - Elle
# Added removal of Installer from /tmp/

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# Version Control 20170111 - Elle
# Adding notification popups

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# Version Control 20170412 - Elle
# Adding codes to quit applications

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

pathToScript=$0
pathToPackage=$1
targetLocation=$2
targetVolume=$3

# Set CocoaDialog Location
CD="/usr/local/XWSUtils/CocoaDialog.app/Contents/MacOS/CocoaDialog"
O365PKG="/tmp/Microsoft_Office_2016_15.32.17030901_Installer.pkg"

rv=`$CD ok-msgbox --no-cancel --title "XWS Mac Notification" --text "Office365 Installation Info" \
--informative-text "Installation will force start in 30 seconds. Please close Microsoft Apps and click on OK to start the installation process now." \
--no-newline --float`


sleep 30

#MSO_PID = MS Outlook
#MSW_PID = MS Word
#MSP_PID = MS PowerPoint
#MSE_PID = MS Excel
#MSN_PID = MS OneNote


MSO_PID=$(ps ax | grep "/Applications/Microsoft Outlook.app/Contents/MacOS/Microsoft Outlook")
MSW_PID=$(ps ax | grep "/Applications/Microsoft Word.app/Contents/MacOS/Microsoft Word")
MSP_PID=$(ps ax | grep "/Applications/Microsoft PowerPoint.app/Contents/MacOS/Microsoft PowerPoint")
MSE_PID=$(ps ax | grep "/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel")
MSN_PID=$(ps ax | grep "/Applications/Microsoft OneNote.app/Contents/MacOS/Microsoft OneNote")

# Shoot them all

kill -9 ${MSO_PID} 2>/dev/null
sleep 2
kill -9 ${MSW_PID} 2>/dev/null
sleep 2
kill -9 ${MSP_PID} 2>/dev/null
sleep 2
kill -9 ${MSE_PID} 2>/dev/null
sleep 2
kill -9 ${MSN_PID} 2>/dev/null


$CD bubble --debug --title "Office365 Notification" --text "Starting installation now and it will complete in 5 mins :) " \
   --background-top "e0e0e0" --background-bottom "f8f8f8" \
   --icon "info"

sleep 10


echo “Removing Old MS Office Apps”
rm -R /Applications/Microsoft\ Excel.app/
rm -R /Applications/Microsoft\ Word.app/
rm -R /Applications/Microsoft\ Word.app/
rm -R /Applications/Microsoft\ PowerPoint.app/
rm -R /Applications/Microsoft\ OneNote.app/

echo “Old MS Office Apps uninstallation is complete”
launchctl unload /Library/LaunchDaemons/com.microsoft.office.licensingV2.helper.plist > /dev/null 2>&1
echo “Killed LaunchD”
rm /Library/LaunchDaemons/com.microsoft.office.licensingV2.helper.plist

rm /Library/Preferences/com.microsoft.office.licensingV2.plist

rm /Library/PrivilegedHelperTools/com.microsoft.office.licensingV2.helper

echo “Killed Licensing Files of Previous version”

echo “Installing new Version of Office 365”

$CD bubble --debug --title "Office365 Notification" --text "Office 365 installation is under progress..." \
   --background-top "e0e0e0" --background-bottom "f8f8f8" \
   --icon "info"

# Kept the installer package name original for records.
installer -pkg $O365PKG -target /
echo “Sleeping 45seconds”
sleep 45

$CD bubble --debug --title "Office365 Notification" --text "Few more seconds, doing the background work..." \
   --background-top "e0e0e0" --background-bottom "f8f8f8" \
   --icon "info"
echo “Adding Dock icons”
/usr/local/XWSUtils/dockutil --add /Applications/Microsoft\ Outlook.app --allhomes --no-restart '/System/Library/User Template/English.lproj'
sleep 5
/usr/local/XWSUtils/dockutil --add /Applications/Microsoft\ Word.app --allhomes --no-restart '/System/Library/User Template/English.lproj'
sleep 5
/usr/local/XWSUtils/dockutil --add /Applications/Microsoft\ Excel.app --allhomes --no-restart '/System/Library/User Template/English.lproj'
sleep 5
/usr/local/XWSUtils/dockutil --add /Applications/Microsoft\ PowerPoint.app --allhomes  --no-restart '/System/Library/User Template/English.lproj'
sleep 5
/usr/local/XWSUtils/dockutil --add /Applications/Microsoft\ OneNote.app --allhomes '/System/Library/User Template/English.lproj'
sleep 5
echo “Getting rid of stupid MAU”
MAU=/Library/Application\ Support/Microsoft/MAU2.0/Microsoft\ AutoUpdate.app
rm -R "${MAU}"

echo “MAU sent out of Galaxy“

rm $O365PKG

$CD bubble --debug --title "Office365 Notification" --text "Installation Complete. Continue using Office365 now." \
   --background-top "e0e0e0" --background-bottom "f8f8f8" \
   --icon "info"
echo “Launch Outlook App“
/usr/bin/open /Applications/Microsoft\ Outlook.app/ 

echo “Installation and cleaning completed.“


exit 0		## Success
exit 1		## Failure