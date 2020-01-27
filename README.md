# SmartGrid
SmartGrid PCF control allows to create a new record directly from the subgrid without navigating to another screen.


Steps
1. Import solution 
2. Add control to subgrid
3. Give required parameters.

![alt text](https://github.com/nijos/SmartGrid/blob/master/SmartGridParameters.JPG)

* Primary Lookup:  Logical name of the lookup field for the relationship.
for example contact subgrid in the account is "*parentcustomerid_account", 

* Primary Entity Set:  entity set name of the current entity where the subgrid is added . 

This is to set the related lookup using web api, you can use Rest builder to get these parameter values correctly.
for example if contact subgrid in account form is using parent customer relationship, to set the account lookup in contact following code is used
var entity = {};
entity["parentcustomerid_account@odata.bind"] = "/accounts(xxxxx-xxxx-xxxx)";

![alt text](https://github.com/nijos/SmartGrid/blob/master/smartgridgif.gif)

Known bugs.
* Does not supporrt N:N
* Validation for empty rows.
* Does not support composite(customer lookups).
* Poor CSS :)

Planned enahancements
* Inline editing.
* Validation for mandatory fields

If this helped you, consider supporting my PCF freebies [![Support via PayPal](https://cdn.rawgit.com/twolfson/paypal-github-button/1.0.0/dist/button.svg)](https://paypal.me/nijojosephraju?locale.x=en_GB)


