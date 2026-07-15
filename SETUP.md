# Antidote dashboard: getting it online and changing it later

## What this is

One folder that becomes one public URL. The page reads the tracker
spreadsheet, recomputes every number and chart, and shows the deck.
Nothing on the page is typed in by hand.

## Files that must be in the folder

- app.py                                    the page itself
- antidote_analysis.py                      the math engine
- Summer_Work_-_Camilo_-_Project.xlsx       the tracker (the only file you edit regularly)
- requirements.txt                          tells the host what to install
- .gitignore                                housekeeping
- Antidote_Review_and_Fluke_Case_Study.pptx the deck, for the download button
- slides/                                   the 11 slide images

The chart images (c_*.png) are redrawn automatically from the tracker,
so you never edit those.

## Going live, one time only

1. Make a free account at github.com.
2. Click New repository. Name it something like antidote-dashboard.
   Make it Public. (Your real dollar figures will be visible. You said
   yes to this on purpose.)
3. Inside the new repository, click Add file, then Upload files.
   Drag in everything listed above, including the slides folder.
   Click Commit changes.
4. Go to share.streamlit.io and sign in with your GitHub account.
5. Click New app. Pick your repository, branch main, file app.py.
   Click Deploy.
6. In about two minutes you get a public link that ends in
   .streamlit.app. That link is what goes on your resume. You can
   change the first part of the address under the app's Settings.

## Changing things after it is live

New sales. Open the tracker on your computer, add the rows the same
way you always do, save. Then in GitHub: Add file, Upload files, drag
the tracker in, Commit. The site redeploys itself and every number and
chart updates. This is the only routine edit.

Deck changes. Edit the .pptx on your computer. Then two uploads to
GitHub: the new .pptx (so the download button serves the new version)
and new slide images in the slides folder. To make the images, export
the deck as PNG from PowerPoint (File, Export), name them slide-01.png
through slide-11.png, or send the .pptx back to me and I will export
and check them.

Dashboard text or layout. That lives in app.py, which is code. Send it
back to me with what you want changed rather than editing it by hand.

## If something breaks

On your app's page, the menu in the bottom right has Manage app. That
shows the logs, which say what went wrong, and has a Reboot button.
Nine times out of ten the cause is a renamed or missing file: the
tracker must keep the exact filename above, or you update INPUT_FILE
at the top of antidote_analysis.py to match.

## The resume line

Antidote LLC, Founder and Operator, with the link:
your-name-antidote.streamlit.app
