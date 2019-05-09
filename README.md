# Utilities

Hey! You've stumbled upon my collection of tools and utilities. Most, hopefully all, are documented on my blog: https://theposhwolf.com.

Most of these scripts are written as functions, so there are a couple of ways to use them:

1. Put them all in a single directory with a module file as I outline in [this post](https://theposhwolf.com/learning_powershell/Toolmaking-in-the-trenches-modules/) and them use ```Import-Module```.
2. Copy and paste the code directly into your PowerShell session, though you'd have to do this every time you wanted access to them.
3. Paste them into your favorite code editor like VS Code or the ISE and then run the script to import the function into your session.
4. Save the file as a .ps1 and use ```Import-Module``` on that file directly.
5. Paste the code into your profile so it loads every time you launch PowerShell.

I prefer #1 so you can have a large number of custom functions and import them all with one command. In fact, you could just put the ```Import-Module``` command in your profile and have it run automatically!
