# PSHTML-AD-Report
My end goal was to create an Active Directory overview report using PowerShell. I looked into PSWinDocumentation but ultimately I wanted the report be interactive. I was looking for basic Active Directory items like Groups, Users, Group Types, Group Policy, etc, but I also wanted items like expiring accounts, users whose passwords will be expiring soon, newly modified AD Objects, and so on. Then I could get this report automatically e-mailed to me daily (or weekly) and I can see what has changed in my environment, and which users I need to make sure change their password soon.  

An overview report like this is also valuable to managed service providers as they can quickly and easily understand a new clients environment, as well as show the customer their own environment. 

While I walk you through the report, you can view it for yourself [here](https://thelazyadministrator.com/wp-content/uploads/2018/12/4-12-2018-ADReport.html)

Below is a screenshot of the Groups tab in the report. Since the report is in HTML you can go to the Active Directory Groups table and search for an item and it will filter the table in real time. If you click the header, "Type" it will order the table by group type instead of name. The pie charts at the bottom can also be interacted with. When you hover over a pie chart it will display the value and count. So if you hovered over the purple portion in Group Membership, it will display "With Members: 18" so I know I have 18 groups that have members.
