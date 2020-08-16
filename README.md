# Kiosk-Slide-Show
These are a couple of addons and a powershell script I built to display and dynamically update a power point slide show 

The powershell script copies the .ppsx file from the server to the host and checks the date modified time on both files and if the server has a newer version 
of the file the script closes powerpoint downloads the new file and starts diplaying the .ppsx again. 

The two addons KioskManager and KioskClient do the following:

KioskManager:

For Kiosk Manager you have three options:
- Expire Slide - Allows you to pick the date, time, and slide number to expire. The expiration time and date are written to the Notes section of that slide
- Unexpire Slide - This removes the date and time from the slide
- Toogle Slide - This toggles the slides hidden property 

![](https://imgur.com/Rpe5Qy5.jpg)

Kiosk Client:

Kiosk Client is a addin that just sits in the background looking at the notes section for each slide. If the time you chose for the current slide
to expire occurs during or before the clide is presented Kiosk Client will mark the slide as hidden and continue to the next slide.
