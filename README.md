# autoMailBot
Usage:
There are four existed sheets, Subject, NameList, BodyList, and ErroList.
In Subject:
1. The second cell in A column is the subject of the mail
2. The second cell in B column is the switch of sending mail automatically. (==1, auto send. ==0 display mail) 
In NameList:
1. The first cell of each column is the title. Input data are from the second row. 
2. Names are stored in A column
3. their mail addresses are stored in B column 
4. Tags are stored in C column
5. Send the mail automatically
In Body:
1. The first cell of each column is the title. Input data are from the second row. 
2. Tags are stored in A column
3.  Body are stored in B column 
4.  SenderDate are stored in C column
5. If "END" appears in a cell of A column, it means the body list is end
Body and Sender and Date in the mail
Auto mail bot:
1. The mail context should be selected from "BodyList" by matching tags. 
2. If there is no matching tag, please generate a new sheet, ErrorList, and copy the wrong row  to the ErrorList. The first cell of each column is the title. Errors are stored from the second row. 
3. The context of mail:
     Hello Name,
     body
     Best,
      SenderDate![image](https://github.com/haoiyang/autoMailBot/assets/129593678/ec6f4d29-25b8-4eb5-b1b7-605aea55a9bd)
