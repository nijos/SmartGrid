# SmartGrid
SmartGrid PCF control allows to create a new record directly from the subgrid without navigating to another screen.


Steps
1. Import solution 
2. add control to subgrid
3. give required parameters.
*Primary Lookup Logical name of the lookup field for the relationship.
for example contact subgrid in the account is #parentcustomerid_account, this is to set the lookup using web api, you can use rest builder to this parameter value correctly.
*Primary Entity Set entityset name of the current entity where the subgrid is added. 

![alt text](https://github.com/nijos/SmartGrid/blob/master/smartgridgif.gif)

Known bugs.
* Does not supporrt N:N
* Validation for empty rows.
* does not support composite(customer lookups).
* Poor CSS :)

Planned enahancements
* Inline editing.
* Validation for mandatory fields

If this helped you, consider supporting my PCF freebies [![Support via PayPal](https://cdn.rawgit.com/twolfson/paypal-github-button/1.0.0/dist/button.svg)](https://paypal.me/nijojosephraju?locale.x=en_GB)


