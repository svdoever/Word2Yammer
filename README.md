# About Word2Yammer

Yammer is a communication tool from Microsoft to connect people and teams accross an organization. See https://www.yammer.com/.

Since Yammer is now also able to edit a post, I have become a bit more enthusiastic about Yammer. Unfortunately, the formatting of text in Yammer is still completely absent, and Yammer has the annoying habit of removing spaces at the beginning of a line, so source‑code with indentation becomes unreadable.

When writing larger posts for Yammer it also a nightmare to do the editing in the "Share something with this group..." box.

Microsoft Word and some PowerShell scripting to the rescue!

I created a PowerShell script to convert a Word document to a plain‑text UTF‑8 document with spaces converted to UTF‑8 non‑breaking spaces. If you open the resulting text file into Notepad and copy/paste it into a Yammer post you will get text with all the special characters used in your Word document with the correct indentation for source‑code.

You can now leverage all the power of Word like spell‑checking, text formatting with numbered items, bullet points and indented source‑code and convert it to a readable Yammer document.

# Installation

1. Create a ```YammerPosts``` folder in your ```Documents``` folder:
   ![Create YammerPosts folder](images/CreateYammerPostsFolder.png)
2. Start a PowerShell console window by typing ```Powershell``` in the explorer bar and press ```Enter```
   ![Start PowerShell console](images/StartPowerShellConsole.png)
3. Run ```Get-ExecutionPolicy```. If it returns ```Restricted```, then run ```Set-ExecutionPolicy AllSigned``` or ```Set-ExecutionPolicy Bypass```
4. Now run the following command to install the ```Word2Yammer.ps``` script in the ```YammerPosts``` folder:
   ```
   (New-Object System.Net.WebClient).DownloadString('https://rawgit.com/svdoever/Word2Yammer/master/Word2Yammer.ps1') > Word2Yammer.ps1
   ```

# Usage 

From the PowerShell console run the script with:

```
 .\Word2Yammer.ps1 ‑Path MyWordsToTheWorld.docx
```

which will result in a file ```MyWordsToTheWorld.txt``` which can be opened in Notepad and the content can be copied to the Yammer text box.

To check for updates run the script with:

```
.\Word2Yammer.ps1 ‑Version
```

# Tips

Some editing tips:

1. If you want an empty line between lines, press <ENTER> twice
2. Copy your source‑code directly from Visual Studio or Visual Studio to your Word document including colored formatting (will be stripped)
3. Create a folder "YammerPosts"  and add the Word2Yammer.ps1 script to this same folder. You can now always go back to your original Word document to edit a Yammer post for republication (by editing the existing Yammer post!)

Happy Yammering!
