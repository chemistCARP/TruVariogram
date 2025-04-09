# TruVariogram

## Description
TruVariogram is an Excel VBA program developed by Christian Romero Prieto as part of work conducted at Noster Tec. This tool is designed to perform variography-based analysis by automatically generating spreadsheets for environmental parameters measured across a series of sampling points along a linear sequence. While the concept of "linear" is straightforward for rivers, in soil analysis, a continuous linear sequence must be established to apply variography effectively. Since evaluations are performed sequentially, this methodology can be applied not only to rivers but also to deep sampling scenarios, such as geological boreholes, vertical sampling in wells, lakes, seas, and similar environments. Additionally, it is suitable for time-based measurements at a fixed point, including those used in the mining-metallurgical industry or fixed sampling stations in water streams where samples are collected at defined intervals.

### Key Features
- **Critical Points and Parameters:** Identifies and evaluates critical points and parameters based on variography.
- **Automated Spreadsheets:** Creates individual worksheets for each environmental parameter measured at sampling points along a line.
- **Variogram Creation:** Generates variograms automatically and includes evaluation tools that compare values against a **Value for Critical Comparisons (VCC)**, a methodology developed to assess critical thresholds.
- **Physical Space Conversion:** Transforms statistical variogram spaces into physical spaces, with results highlighted in graphical outputs.
- **Joint Variograms:** Introduces a novel concept of combined variograms, allowing users to create a single chart with selected variograms.
- **Data Cloning:** Supports parallel evaluations by cloning datasets.
- **ChartMaps (Map Integration):** If a map (e.g., from Google Earth) with sampling point coordinates is provided, TruVariogram generates a custom Excel chart ("ChartMaps") where critical points and parameters are displayed as circles. Circle sizes scale based on the number of critical parameters (impacts) detected.
- **Sampling Prediction:** Enables users to predict the required number of sampling points based on desired or achievable sampling error.
- **Spreadsheet Organization Tools:** Includes features to organize worksheets with color differentiation and naming conventions for ease of use.

This program was developed to support environmental analysis and is shared as part of a scientific contribution by Noster Tec.

## Requirements
- Microsoft Excel (version 2016 or later recommended).
- Macros must be enabled to run the VBA code.

## Installation and Usage
1. Download the file `TruVariogram.xlsm` from this repository.
2. Open the file in Microsoft Excel and enable macros when prompted.
3. Input your data:
   - Ensure sampling points are organized in a linear sequence (e.g., along a river or a defined soil transect).
   - Include environmental parameter measurements in the designated input sheet.
4. Run the program:
   - Access the main interface (e.g., a button or menu) to generate variograms and analysis outputs.
   - Follow on-screen prompts to select parameters for joint variograms or ChartMaps (if using a map).
5. Review results in the automatically generated worksheets and charts.

For detailed instructions, refer to comments within the VBA code or contact the author.

## Files Included
- `TruVariogram.xlsm`: The main Excel VBA program.

## License
Copyright 2025 Christian Romero Prieto  

This project is licensed under the **Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)**.  
- **Permissions:** You are free to share, copy, and adapt this code provided you give appropriate credit to the author (Christian Romero Prieto) and use it solely for non-commercial purposes.
- **Restrictions:** Commercial use of this code is prohibited without explicit written permission from the author.
- See the [LICENSE](LICENSE) file or visit [https://creativecommons.org/licenses/by-nc/4.0/](https://creativecommons.org/licenses/by-nc/4.0/) for full details.  
- For commercial use inquiries or additional permissions, contact: [your email].

## Author
- **Christian Romero Prieto**  
- Developed under **Noster Tec**  
- Contact: ductor@nostertec.com

## Citation
If you use TruVariogram in your research, please cite it as:  
Romero Prieto, Christian, "TruVariogram: A VBA Tool for Variography-Based Environmental Analysis," developed at Noster Tec, 2025. Available at https://github.com/chemistCARP/TruVariogram/.

## Acknowledgments
This tool was created as part of environmental research efforts at Noster Tec in Bolivia. Special thanks to mentors: Carlos Arze Landivar, Cesin Curi Sabja, Juan Cristobal Birbuet.

## Notes
- Ensure your sampling data follows a linear sequence for optimal performance.
- For map-based ChartMaps, provide coordinates in a compatible format (e.g., latitude/longitude from Google Earth).

---
