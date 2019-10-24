# Excel Visual Basic for Applications Pharmacy Functions Library

## Install Instructions:
Code library add-in for custom pharmacy functions in Excel. To install, download 'RxFxLibrary.xlam', click the 'Developer' tab, click 'Excel Add-ins', click 'Browse...', and find the location of the file. The custom menu will be located under the 'Formulas' tab. If you are using this for clinical practice, please make sure you are using the correct units (metric vs US). All functions have been tested against multiple online calculators and hand calculations. Please use the functions at your own risk.

If you have any questions or find any bugs, please contact EszopiCoder at pharm.coder@gmail.com

## List of Current Functions:

### Height and Weight
#### Body Mass Index (BMI)
**Function**
- `RxCalc_BMI()`


**Equation**
- BMI = Weight / Height^2


**Parameters**
`Height` String
Required. Height of person in inches or centimeters. Height in inches may be formatted as 5'10".
`Weight` Single
Required. Weight of person in pounds or kilograms.
`Metric` Boolean
Optional. Measurement units of height and weight. True=Metric (Default); False=US
**Returns**
Variant
BMI in kg/m^2 of person given height and weight.
#### Body Surface Area (BSA)
**Functions**
- `RxCalc_BSA_DuBois()`
- `RxCalc_BSA_Mosteller()`


**Equations**
Du Bois Formula:
- BSA = 0.007184 * Weight^0.425 * Height^0.725


Mosteller Formula:
- BSA = (Height * Weight)^0.5 / 60


**Parameters**
`Height` String


Required. Height of person in inches or centimeters. Height in inches may be formatted as 5'10".


`Weight` Single


Required. Weight of person in pounds or kilograms.


`Metric` Boolean


Optional. Measurement units of height and weight. True=Metric (Default); False=US


**Returns**
Variant


BSA in meters^2 of the person given height and weight.
#### Ideal Body Weight; over 60 inches (IBW)
**Function**
- `RxCalc_IBW()`


**Equations**
Devine Formula:
- IBW (Male) = 50kg + 2.3kg for each inch above 60 inches
- IBW (Female) = 45.5kg + 2.3kg for each inch above 60 inches


**Parameters**
`Height` String
Required. Height of person in inches or centimeters. Height in inches may be formatted as 5'10".
`Female` Boolean
Required. Sex of the person. True=Female; False=Male
`Metric` Boolean
Optional. Measurement units of height. True=Metric (Default); False=US
**Returns**
Variant
IBW in kilograms of person given height and sex.
#### Adjusted Body Weight; obese (AdjBW)
**Function**
- `RxCalc_AdjBW()`


**Equations**
Devine Formula:
- AdjBW = IBW + 0.4*(Actual Body Weight - IBW)


**Parameters**
`Height` String
Required. Height of person in inches or centimeters. Height in inches may be formatted as 5'10".
`Weight` Single
Required. Weight of person in pounds or kilograms.
`Female` Boolean
Required. Sex of the person. True=Female; False=Male
`Metric` Boolean
Optional. Measurement units of height and weight. True=Metric (Default); False=US
**Returns**
Variant
AdjBW in kilograms of person given height, weight, and sex. Only use for obese patients.
#### Ideal Body Weight; under 60 inches (IBW)
**Functions**
- `RxCalc_IBW_Intuitive()`
- `RxCalc_IBW_Baseline()`
- `RxCalc_IBW_Hume()`


**Equations**
Intuitive Formula:
- IBW (Male) = 50kg - 2.3kg for each inch below 60 inches
- IBW (Female) = 45.5kg - 2.3kg for each inch below 60 inches


Baseline Formula:
- IBW (Male) = 50kg - 0.833kg for each inch below 60 inches
- IBW (Female) = 45.5kg - 0.758kg for each inch below 60 inches


Hume Method:
- IBW (Male) = (0.3281 x Weight in kg) + (0.33939 x Height in cm) - 29.5336
- IBW (Female) = (0.29569 x Weight in kg) + (0.41813 x Height in cm) - 43.2933


**Parameters**
`Height` String
Required. Height of person in inches or centimeters. Height in inches may be formatted as 5'10".
`Weight` Single
Required for Hume method only. Weight of person in pounds or kilograms.
`Female` Boolean
Required. Sex of the person. True=Female; False=Male
`Metric` Boolean
Optional. Measurement units of height and weight. True=Metric (Default); False=US
**Returns**
Variant
IBW in kilograms of person given height and sex. Weight required for Hume method.

### Renal
#### Cockcroft-Gault Creatinine Clearance (CrCl)
**Function**
- `RxCalc_CrCl()`


**Equation**
CrCl = ((140 - Age) * Weight) / (72 * sCr)
**Parameters**
`Age` Byte
Required. Age of person in years.
`Weight` Single
Required. Weight of person in pounds or kilograms.
`sCr` Single
Required. Serum creatinine in mg/dL.
`Female` Boolean
Required. Sex of the person. True=Female; False=Male
`Metric` Boolean
Optional. Measurement units of weight. True=Metric (Default); False=US
**Returns**
Variant
CrCl in mL/min of person given age, weight, serum creatinine, and sex.
#### Modification of Diet and Renal Disease Study (MDRD)
**Function**
- `RxCalc_GFR_MDRD()`


**Equation (4-variable)**
eGFR = 175 * sCr^-1.154 * Age^-0.203 * 0.742 (if female) * 1.212 (if black)
**Parameters**
`Age` Byte
Required. Age of person in years.
`sCr` Single
Required. Serum creatinine in mg/dL.
`Female` Boolean
Required. Sex of the person. True=Female; False=Male
`Black` Boolean
Optional. Race of the person. True=Black (Default); False=Others
**Returns**
Variant
eGFR in mL/min/1.73m^2 of person given age, serum creatinine, sex, and race.
#### Chronic Kidney Disease Epidemiology Collaboration (CKDEPI)
**Function**
- `RxCalc_GFR_CKDEPI()`
**Equation (4-variable)**
eGFR = 141 * min(sCr/k, 1)^a * max(sCr/k, 1)^-1.209 * 0.993^Age * 1.018 (if female) * 1.159 (if Black)
**Parameters**
`Age` Byte
Required. Age of person in years.
`sCr` Single
Required. Serum creatinine in mg/dL.
`Female` Boolean
Required. Sex of the person. True=Female; False=Male
`Black` Boolean
Optional. Race of the person. True=Black (Default); False=Others
**Returns**
Variant
eGFR in mL/min/1.73m^2 of person given age, serum creatinine, sex, and race.

### Diabetes
#### Correction factor dosing
**Function**
- `RxCalc_CorrectionFactor()`


**Equation**
Insulin Sensitivity (IS):
- IS (Rapid insulin) = 1800 / Total Daily Dose
- IS (Regular insulin) = 1500 / Total Daily Dose
Correction Factor Dosing:
- CF = (Actual Blood Glucose - Target Blood Glucose) / IS
**Parameters**
`TDD` Integer
Required. Total daily dose of insulin (basal+bolus) in units.
`ActualBG` Integer
Required. Actual blood glucose in mg/dL.
`TargetBG` Integer
Required. Target blood glucose in mg/dL.
`RapidIns` Boolean
Optional. Type of bolus insulin used for meals. True=Rapid (Default); False=Regular
**Returns**
Variant
Insulin dose in units. Add to set dose or carb counting dose.
#### Carbohydrate counting dosing
**Function**
- `RxCalc_CarbCounting()`


**Equation**
Carb:Insulin Ratio (C:I): 
- C:I = 500 / Total Daily Dose
Carb Counting Dose (CC):
- CC = Carbs / C:I
**Parameters**
`TDD` Integer
Required. Total daily dose of insulin (basal+bolus) in units.
`Carbs` Integer
Required. Carbohydrates in grams.
**Returns**
Variant
Insulin dose in units. Add to correction factor dose.

### Kinetics (Coming soon!)
- Vancomycin dosing
- Aminoglycoside dosing
