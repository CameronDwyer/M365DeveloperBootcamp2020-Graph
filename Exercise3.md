# Exercise 3: Warehouse Packing App with Microsoft Graph Toolkit

 * [Exercise 1: Lab setup](Exercise1.md)
 * [Exercise 2: Build a JavaScript single-page app with Microsoft Graph](Exercise2.md)
 * [Exercise 3: Warehouse Packing App with Microsoft Graph Toolkit](Exercise3.md) **(You are here)**
 * [Resources](Resources.md)
 
The scenario for this exercise is that our company sells products online and those online orders are created in a SharePoint List. We are going to build a lightweight single page web application using HTML, CSS and JavaScript for the warehouse packing staff to be able to view the orders that are ready for packing on a small device. The warehouse staff will carry this device with them as move about the warehouse collecting and packing the products into a box for the customer. When they are done our web application will also allow the packing staff to mark the order as packed.

## Clean up after completing Exercise 2
If you still have a command-line running your local http-server from Excerise 2 you will need to stop if before continuing as starting another http-server will create it on a different port number. To stop the http-server close your command-line window or use CTRL+C.

## Create the Warehouse Order list in SharePoint
Before we can start building our web application we need to setup the scenario by creating an `Orders` list in SharePoint. 

1. Create a new SharePoint site (select Team site as the type and give it the name `Warehouse`)

![opensharepoint](./images/warehouse-02.jpg)

2. Add a new List to the SharePoint site

![newlist](./images/warehouse-03.jpg)

* Choose `Blank` rather than a template 
* Name = `Orders`
* Check the box to `Show in Site Navigation` (so we can find our list easily)

![newlist](./images/warehouse-04.jpg)

3. Add the following columns to the `Orders` list

When adding columns, ensure you use the modern UI as shown below. This will allow you to have spaces in your column names and generate an internal column name we use in code that doesn't have the space in it .

![addcolumn](./images/warehouse-05.jpg)

* Be very careful of column names as the code snippets reference these columns

| Column Name    | Column Type             | Notes                                                     |
|----------------|-------------------------|-----------------------------------------------------------|
| Title          | Single line of text     |                                                           |
| Status         | Choice (Drop-Down Menu) | Allow choices of: Requested, Ready for Packing, Delivered (enter each on a separate line) |
| Address        | Single line of text     |                                                           |
| Total Items    | Number                  |                                                           |
| Order Manifest | Multiple lines of text  |                                                           |
| Contact Name   | Single line of text     |                                                           |


4. Rename the `Title` column to `Order Number` by selecting the column header | Column Settings | Rename

![newlist](./images/warehouse-06.jpg)


5. Populate the orders list with 3 orders
* Give 2 orders the `Status` of `Ready for Packing`
* Give 1 order the `Status` of `Requested`

Your populated list should look similar to the following.

![orders](./images/warehouse-07.jpg)

6. Create a new list index for the `Status` column
 * We need to do this because we want to query from our web application to find orders in a `Status` of `Ready for Packing`
 * Create the index under List Settings | Columns | Indexed Columns

![orders](./images/warehouse-08.jpg)

![orders](./images/warehouse-09.jpg)

![orders](./images/warehouse-10.jpg)

🎉Celebrate a little, SharePoint is setup and ready to go. Let's go develop our web application! 😎

## Create project directory
Create a new directory for the Warehouse Packaging App

## Start a local web server
In this section you will use http-server to run a simple HTTP server from the command line.

1. Open your command-line interface (CLI) in the directory you created for the project.

2. Run the following command to start a web server in that directory.
```
npx http-server -c-1
```

3. Open your browser and browse to `http://localhost:8080`.

![localserver](./images/warehouse-01-server.jpg)

## Create index.html
Create a new file in the project directory called `index.html` with the following content.
```
<html>
  <head>
    <link rel="stylesheet" type="text/css" href="styles.css">
  </head>
  <body>
    <header>
      <ul>
        <li>Warehouse App</li>
      </ul>
    </header>
  </body>
</html>
```

## Create a stylesheet for the application
Create a new file in the project directory called `styles.css` with the following content.
```
body {
    font-family: 'Segoe UI';
    font-size: 12px;
    background-color: #F3F2F1;
    margin: 0px;
}

ul {
    list-style-type: none;
    margin: 0;
    padding: 0;
    overflow: hidden;
    background-color: #333;
  }
  
li {
    font-size: 18px;
    float: left;
    display: block;
    color: white;
    text-align: center;
    padding: 14px 16px;
    text-decoration: none;
}

.order {
    display: flex;
    flex-direction: column;
    color: #444;
    border: 1px #EDEBE9 solid;
    width: 270px;
    height: 200px;
    background-color: #ffffff;
    margin: 10px;
    float:left;
    box-shadow: 2px 2px #E4E3E2;
    border-radius: 10px;
}

.orderheader {
    font-weight: bold;
    font-size: larger;
    background-color: rgb(200, 203, 233);
    padding: 5px;
    height: 25px;
    border-radius: 10px 10px 0 0;
}

button {
    text-align: center;
    padding: 10 5 10 5;
    font-family: 'Segoe UI';
    font-size: 14px;
    background-color: rgb(178, 207, 178);
    border-radius: 0 0 10px 10px;
}

button.inactive {
    visibility: hidden;
}

.orderbody {
    flex-grow: 2;
    white-space: pre-line;
    margin: 10px;
}

.orderfooter {
    font-weight: bold;
    font-size: larger;
    padding: 10px;
    height: 60px;
}
```

## Add reference to Microsoft Graph Toolkit (MGT)

## Create App registration in Azure Active Directory
TODO
Link to Beth's article


## Add mgt-msal-provider for authentication

## Add mgt-login component and customise the styling


Toolkit component allow for styling customisation. [Read this article to learn how it works](https://developer.microsoft.com/en-us/graph/blogs/a-lap-around-microsoft-graph-toolkit-day-4-customizing-components/)

Add the following style rule to `styles.css`
```
mgt-login {
    --color: white;
    --padding: 0 20px;
}
```

## Use Graph Explorer to verify query for SharePoint items
TODO

## Add mgt-get component to fetch SharePoint items
TODO

## Update SharePoint item status using Graph SDK (obtained through MGT Provider)
TODO