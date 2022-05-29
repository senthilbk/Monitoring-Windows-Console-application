# Monitoring-Windows-Console-application

Objective of Python Script 
==========================


One of my Client approached me with a problem that They have Legacy windows based Application which is running on mutiple windows server and also they have configured in Vmware load balancer 

The Loadbalancer will distribute request across the server in which Application is running ,For some times the Application got hanged and could not process the request but it seems that the load balancer is not aware of that.

During the Hanging the application, they used to get pop-up which requires manual intervention to procced further 


So my client wanted that they need to know when application got hanged and mail needs to be triggered and the pop-up and hanged application  needs to be killed, so that application can be started to have zero downtime  


What Does this python Script do?
==============================

This python script will get list of visible console windows on server and Cchek if the Client Windows Legacy console Application is running/visible , then The script will not do anything ,

But If the Pop up is visible , then POP window and legacy application will be killed 


Input for this Script 
====================
 serviceName='HDSentinel.exe' -->Line NO: 143--The windows  service of the application
 
 consoleName='Hard Disk Sentinel'---> Line NO: 144==The Title of Leagcy windows based console Application 
 
 consoleNameDll='Advanced Power Management'-->Line 146==The Title of POPUP because of hanging 
 
