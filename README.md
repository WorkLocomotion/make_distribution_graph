# Work Locomotion — make_distribution_graph

Python scripts for generating workforce OVP distribution data and creating the smooth, shaded curve used in the Corporate OVP Distribution Graph.

This repository completes Part 3 of the Corporate OVP Analysis series, showing how opportunities for meaningful work are distributed across company occupations.

---

## Overview

### build_ovp_histogram.py
- Reads your Step 2 output file: Company Job Titles – Mapped.with_work_values.xlsx  
- Computes histogram data across the six OVP ranges  
- Exports a new Excel file: OVP_Distribution_Data.xlsx  
- Provides the data foundation for Excel visualization  

### build_ovp_curve.py
- Reads the histogram data  
- Generates a smooth, shaded curve (Catmull–Rom or cubic Bézier)  
- Outputs a preview plot (ovp_distribution_curve.png) and Excel-ready data  

Together, these scripts transform occupational work values into a continuous distribution that reveals how “opportunity for meaningful work” is spread across your workforce.

---

## Requirements

If you are running locally, install the required packages:

```bash
pip install pandas numpy scipy matplotlib openpyxl

How to Use
1. Ensure your Step 2 file exists: Company Job Titles – Mapped.with_work_values.xlsx
2. Run the histogram script: python build_ovp_histogram.py
3. Run the curve script: python build_ovp_curve.py
4. Open the resulting Excel file and follow the distribution-graph tutorial on Substack: https://worklocomotion.substack.com/

Background
This repository continues the Work Locomotion analytical sequence:
Part 1: Map company job titles to O*NET SOC codes
Part 2: Enrich titles with O*NET Work Values
Part 3: Visualize the distribution of meaningful work opportunities
The resulting graph illustrates how well your workforce’s roles align with the psychological needs and intrinsic motivations embedded in their work.

Released under the MIT License.

WORK LOCOMOTION: Make Potential Actual
