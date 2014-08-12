Student Handin Project Member Refresh
=====================================

Student Handin Project Creator generated student hand in folders for given course assignments, automatically giving each student full rights to their respective folders. The problem is that students are always registering or deregistering from the course at any given time, meaning that the possibility of a student's personal handin folder not existing is very real.


Student Handin Project Member Refresh runs every 10 minutes, checking for new group members and generating their personal handin folders as required.

Note: Created for Commerce I.T.
Note: Update: Handin shifted from a Novell server to a Windows 2003 server.

Created by Craig Lotter, April 2006

*********************************

Project Details:

Coded in Visual Basic .NET using Visual Studio .NET 2005
Implements concepts such as Folder manipulation, Shell scripting.
Level of Complexity: Simple

*********************************

Update 20070404.03:

- With the decommissioning of the Novell based Comlab file server, it was decided to move the handin functionality to the online sphere, namely the Commerce webserver. This is a Windows 2003 server so the Project Creator application needed to be updated accordingly.
- Now allows the user to specify creating student folders for a Novell or Windows environment. (Note: Windows environment - folder rights are determined by the parent folder rights.
