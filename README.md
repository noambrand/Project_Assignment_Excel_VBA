# Project_Assignment_Excel_VBA
Automatically create personalized tables of assignments for each employee, allowing them to focus on only their assignments and priorities and not be distracted by other employee assignments that do not concern them. Create management reports with all projects and employees together.

How does it work: It combines different sheets with the same fields to one table, copies the combined table to a new sheet and
Filters it by employee name to separate sheets, AutoFits the table nicely and finally saves the sheets to separate excel files
in a new folder that has the name of the workbook.
The AutoFilter in the code is in columns E,F (See Fields 5,6 in the AutoFilter)

Macro steps:
1) Combines Number of first sheets with the same fields to one sheet.
2) Makes copies the combined sheet to the end of the workbook.
3) Rename sheets by each employee name.
4) Filter each sheet by employee name +not yet complite assignments.
5) Each sheet still has all of the assignments(the user can cancel filter by autofilter).
6) Autofit columns width and sheet from right to left
7) Split the workbook sheets to separate files in a folder.
   The file name will be the sheet name + date of creation.

The example has 4 sheets- 3 big projects that have many assignments and 1 sheet that has 4 small projects with few assignments.
Ps. If the assignment is e.g. "Jerry+Noam" it will be present in Jerry's table and also Noam's table.

Question: Why not create a table with all the projects and assignments in the first place that will save the need to combine different sheets?
Answer: When your table gets filled with projects that have many assignments it becomes difficult to manage efficiently, the table gets messy and editing is more complicated with autofilter on,  you have to keep filtering the table which is time-consuming and prone to errors.

מטרת הקובץ: עבור מנהלי פרויקטים/ מנהלי מוצר שמנהלים פרויקטים בסביבת אקסל בסביבה מרובת פרויקטים ועובדים. 
יצירה אוטומטית של טבלת משימות מותאמת אישית עבור כל עובד, המאפשרת לעובד להתמקד רק במטלות ובסדרי העדיפויות שלו.
כמו כן יצירת דוח אוטומטי מרכז להנהלה שכולל את כל הפרויקטים והעובדים.

איך זה עובד: המאקרו משלב גיליונות שונים עם אותם שדות לטבלה אחת, מעתיק את הטבלה המשולבת לגיליון חדש ומסנן אותה לפי שם עובד 
לגיליונות נפרדים, מתאים אוטומטית את העמודות לטבלה יפה ולבסוף שומר את הגיליונות לקובצי אקסל נפרדים בתיקייה חדשה עם השם של חוברת העבודה.
הסינון האוטומטי בקוד נמצא בעמודות E,F  (ראה שדות 5,6 במסנן האוטומטי)

הדוגמה כוללת 4 גיליונות: 3 פרויקטים גדולים שיש להם הרבה מטלות וגיליון אחד שיש לו 4 פרויקטים קטנים עם מעט מטלות.
נ.ב. אם המשימה היא למשל של  "מוטי+נועם" המשימה תהיה קיימת גם בטבלה של מוטי וגם בטבלה של נועם.
נ.ב 2: הערכים בטבלאות לדוגמא שובשו באופן מכוון לשם שמירה על פרטיות.

שלבי מאקרו:
1) משלב מספר גיליונות ראשונים עם אותם שדות לגיליון אחד.
2) יוצר העתקים של הגיליון המשולב בסוף חוברת העבודה.
3) משנה את שם הגיליונות לפי שם העובד.
4) מסנן כל גיליון לפי שם העובד והמטלות שטרם הושלמו.
5) בכל גיליון עדיין יש את כל המשימות (המשתמש יכול לבטל את הסינון ולצפות במשימות של עובדים אחרים).
6) התאמה אוטומטית של עמודות רוחב וגיליון מימין לשמאל
7) מפצל את גיליונות חוברת העבודה כדי להפריד קבצים בתיקייה.
    שם הקובץ יהיה שם הגיליון + תאריך היצירה.

שאלה: למה לא ליצור טבלה עם כל הפרויקטים והמטלות מלכתחילה שיחסוך את הצורך בשילוב גיליונות שונים?
תשובה: כשטבלה מתמלאת בפרויקטים שיש להם משימות רבות, זה מקשה על ניהול יעיל,  הסינון מסרבל את יכולת העריכה, והצורך לסנן כל פעם מחדש גוזל זמן ורגיש לשגיאות.


