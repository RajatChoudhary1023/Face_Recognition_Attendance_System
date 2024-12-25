Face Recognition Attendance System
<img width="500" alt="Attendance" src="https://github.com/user-attachments/assets/65eb6e38-a1e4-4948-9dde-2870f1973b0f" />

Steps to get started<br>
1.Create a folder named data(This will have the dataset).<br>
2.Run the datagathering.py and collect the dataset.Make sure you give id from 0,1,2... for convinience.<br>
3.Go to app.py in that change the gc,worksheet variables(Search them ctrl+f) with your credentials.Same in index.html change the "YOUR_GOOGLE_SPREADSHEET-LINK" with your link.<br>
4.In app.py search for the variable 'checkface' in that list add 0's as per the number of unique people images(for eg- if there are 3 people then it becomes checkface=[0,0,0].<br>
5.Search for names variable and add the names of the people make sure you follow the order of Id(in datagathering)(for eg. ['ABC','XYZ'] here id=0 is for ABC and id=1 for XYZ).<br>
6.Run the app.py and index.html.
