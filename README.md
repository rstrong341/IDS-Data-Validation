#  IDS Data Validator (VBA)

A VBA script that adds a button to validate data on IDS forms. If incorrect data is detected, the script highlights the issues using conditional formatting and provides guidance on how to correct them.

---

## 📖 Description
This project automates data validation for IDS (Information/Data Sheets) in Excel.  
Key features:
- Adds a validation button to the sheet.
- Checks for incorrect or missing data.
- Highlights errors using conditional formatting.
- Displays instructions for fixing issues.
- 
<img width="350" height="324" alt="Screenshot 2025-10-08 104344" src="https://github.com/user-attachments/assets/576654fb-c38f-4dc4-8309-3c5805eea3c0" />

---
##  Validations
This IDS Iteration checks the below feilds to ensure accuarcey:

- **Required fields are not blank**  
  If a Cell labled as a required field is missing data, the field will be highlighted and explained in the "Instruction to Correct Errors" section 														


- **Pallet Conversions follow requirements**  
  If a cell in this section does not follow the required quantitative logic this check will highlight the cells that need to be corrected and an explanation will be provided in the "Instruction to Correct Errors" section 

- **There is not a metric/imperial mix**  
  If a the dimensions and weights in this section are mixed between imperial and metric this check will flag both cells and an explanation will be provided in the "Instruction to Correct Errors" section 

- **GTINS equal 14 digits**  
  If a Cell in the GTIN section is listed with less or greater than 14 digits this check will flag the impacted cells and an explanation will be provided in the "Instruction to Correct Errors" section 
  
- **Net Weight can't be greater than gross weight**  
  If a Cell in the net weight is greater than the gross weight of any UOM in the weight section of the IDS this check will flag the impacted cells and an explanation will be provided in the "Instruction to Correct Errors" section  
  
- **Pallet size Requirements**  
  If a Cell in the pallet size section on the IDS breaches the maximum size, this check will flag the impacted cells and an explanation will be provided in the "Instruction to Correct Errors" section  
  
- **Case Net/Gross Weight max**  
  If a Cell in the Weights section on the IDS breaches the maximum weight, this check will flag the impacted cells and an explanation will be provided in the "Instruction to Correct Errors" section  



<img width="350" height="324" alt="Screenshot 2025-10-08 104344" src="https://github.com/user-attachments/assets/576654fb-c38f-4dc4-8309-3c5805eea3c0" />
