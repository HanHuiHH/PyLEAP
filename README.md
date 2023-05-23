# PyLEAP - Python Programming for LEAP API

Welcome to PyLEAP - Python Programming for LEAP API. This project is for further research using LEAP software and Python, including automatic analysis and multi-objective optimization.  

In the sample code, we use [SALib](https://github.com/SALib/SALib) to do Monte Carlo simulation and Sobol' global sensitivity analysis.

The code now just have annotation in Chinese.

中文版README请打开**README(中文).md**。


## LEAP and its API

### LEAP software

The Low Emissions Analysis Platform (LEAP) is a widely-used software tool for energy policy, climate change mitigation and air pollution abatement planning developed at the Stockholm Environment Institute (SEI). See more information in the [website of LEAP](https://leap.sei.org). You can also get start quickly by their [Open Youtube Courses](https://www.youtube.com/watch?v=y4b2KCIxOJU&list=PLX-Kjcc7K01EOTxozEEBu2aerJmZ6ZfRT&ab_channel=LEAPPlatform).

### API in LEAP

LEAP can act as a standard "COM Automation Server," meaning that other Windows programs can control LEAP directly: changing data values, calculating results, and exporting them to Excel or other applications. The API even provides functions for examining or changing LEAP's data structures. This ability to program LEAP can be very powerful.   

For example, you could write a short script that could run LEAP calculations many times, each time with a different set of input assumption. LEAP's results could then be output to Excel or processed in the script and used to calculate revised assumptions for subsequent LEAP calculations. In this way LEAP's basic accounting calculations could be coupled with more sophisticated algorithms such as goal-seeking or optimizing algorithms.  

For more information, open the **content** in LEAP software, and see **Advance Topics/Automating LEAP (API)**.  


## Performance

### Iteration

On a system comprising an i7 8700 processor, 16GB of RAM, and Python 3.9.16, each iteration took approximately **5 seconds**. which is relatively slow for a large sample size. To ensure the functionality of the code, we suggest you to initially conduct testing using several small samples. Due to the limit of LEAP software, the multi-processes running method is not avaliable in PyLEAP. We are trying to realize this function.

### Result Processing

After iterations, the results will be saved. If you need a big sample size (running time is more than 10 hours), try to take it apart into several parts, and merge results of these parts. The result processing will not cost long time.


## Code Structure

### Calculate and Check (CalculateAndCheck.py)

This code calls calculate function in LEAP software, and print several key values (such as energy intensity decreasing rate in 5 years) in your console. You can just use this code to check key values in your LEAP area.

### Import from Excel (ImportFromExcel.py)

This code is to import value from excel file in the **ImportExcel folder**. The simulation need to define distribution of key parameters, so that samples can be generated. Mean value and standard variation are the key parameters of normal distribution, and they can be edited in this excel file. 

### Monte Carlo Simulation (MonteCarlo-Simulation.py)
This code integrated former two codes into Monte Carlo simulation by SALib Library. Results are saved in the Result folder.

### Monte Carlo Simulation Analysis (MonteCarlo-Analysis.py)
This code use the results of Monte Carlo to do analysis result distribution, and draw some pictures.

### Sobol' GSA (Sobol'GSA.py)
This code use the results of Monte Carlo to do Sobol' global sensitivity analysis, and draw some pictures.

### Bug Fix By Initialize LEAP (BugFIX.py)
There may occur some errors if you abort running and rerun codes. This code will fix this type of error.
