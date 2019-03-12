# DSL Test Utility

Tool to automate account information retrieval and DSL modem testing. Written in AutoHotkey(AHK)

This software is not useful as-is, it was developed as an internal ISP tool. The script has been scrubbed to remove the actual URLs and tool names used at the company that I worked for.  This repository exists to give an example of my approach to problem solving using software.  

In January of 2019, I became obsessed with automating the boring bits of my support job.  This is my first crack at any kind of programming past basic tutorials and 'hello world' - I realize the structure of the code is probably a mess.  I just first needed it to work.  It took about 3 weeks to get this all working correctly.  

I learned a good deal about a lot of things, namely:

* Scraping DOM elements with Selenium Web Driver
* A few regular expressions to manipulate data
* Splitting data into arrays for processing
* Building a GUI with AutoHotkey
* Patience during the software building process

The DSL support contract that I was on ended before I could get my team to use this software, I had been looking forward to getting bug reports and feature requests.

I left some code sections in that are commented out because I left the contract so quickly that I couldn't remember why they were commented out.  

## Prerequisites

* [Microsoft Windows](https://www.microsoft.com/en-us/windows) - Written and tested on Windows 7
* [AutoHotkey](https://www.autohotkey.com/download/) - Minimum version: 1.1.30.01
* [Selenium Basic version 2.09.0](https://github.com/florentbr/SeleniumBasic/releases/download/v2.0.9.0/SeleniumBasic-2.0.9.0.exe)

## Installing/Configuring

* [Cross browser web scraping with AutoHotkey and Selenium](https://the-automator.com/cross-browser-web-scraping-with-autohotkey-and-selenium/)
* [AutoHotkey help](https://www.autohotkey.com/docs/AutoHotkey.htm)

## Built With

* [SciTE4AutoHotkey](https://github.com/fincs/SciTE4AutoHotkey) - Code editor for AutoHotkey
* [Darchon](https://github.com/ahkon/Darchon) - Excellent dark theme syntax highlighting for the AutoHotkey language

## Author

Brian Watt

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

Many thanks to Joe Glines for providing free AHK resources - https://the-automator.com/
