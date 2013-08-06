# AutoCollapse

## Summary

Visual Studio 2012 add-in that basically collapses everything when you open a file.

## Description

I found that I tend to press `CTRL + M, CTRL + O` (Collapse to Definitions) whenever I opened a `.cs` file, and `CTRL + K, CTRL + O` (Start Automatic Outling) whenever I opened a `.cshtml` file from the Solution Explorer. 

In the past, I could have set up a macro to do this automatically, but this functionality has been removed in Visual Studio 2012; thus I build this add-in.

It basically runs the `Edit -> Outlining -> Start Automatic Outling` and `Edit -> Outlining -> Collapse to Definitions` commands whenever a file is opened. Note that it _only_ does this when files are opened from Solution Explorer, since I did not want everything collapsed when, say, searching or debugging a file.

I threw this together for my own convenience. I do not plan to continue developing this further, unless I personally encounter a bug that irritates me enough to fix it. Use this at your own risk, and feel free to fork the repository.

I also have no idea if it will work on older versions of Visual Studio.