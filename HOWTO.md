# Quickstart with clasp and Git

1.  cd into (or create) your project folder.
2.  clasp login                                       # authenticate with Apps Script
3.  clasp create --type standalone --title "My Script"  # or `clasp clone <scriptId>` if already created in Apps Script
4.  git init
5.  cp .claspignore .gitignore                        # copy ignore rules
6.  git add .
7.  git commit -m "Initial import"
8.  git branch -M main
9.  git remote add origin <your-repo-url>
10. git push -u origin main

11. clasp pull                                       # pull down code from Apps Script
12. clasp push                                       # push local changes to Apps Script

> **Note:** We embed a company logo into the PDF by loading its blob via DriveApp and converting to a base64 data URL. Be sure to set the logo file’s sharing to **Anyone with the link can view** so Apps Script can access it at runtime.