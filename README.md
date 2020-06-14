# rti-web-scrapper
RTI Analytics' web scrapper script using Selenium 

## Requirements
```bash
~ pip3 install -r requirements.txt
```
Install Chrome Webdriver as well. You can refer to this [wiki](https://github.com/SeleniumHQ/selenium/wiki/ChromeDriver) for the installation guide.

## How to Run
Change the star and end value of `iterate` function. The last parameter is the filename.
You can uncomment `runInParallel` and call `iterate` function with different start, end and filename value to run the script in parallel.

```bash
~ python3 scripts.py
```
