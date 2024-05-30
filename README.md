# AutoTimeClock

## Introduction

This project is designed to simplify and automate clock-in and clock-out functionalities within the given shift schedule for a user on Microsoft Teams.

## Getting Started

### Prerequisites

Ensure you have Git installed on your system to clone the repository.

### Cloning the Repository

To get started, clone the project to your local machine using the following command:
`git clone https://github.com/ssa2973/AutoTimeClock`

## Using the Application

After cloning the repository, navigate to the project directory and run `AutoTimeClock.exe`. Optionally, you can also set the application to run at system startup by following these steps:

1. Open Run (Windows+R) and type shell:startup.
2. Create a shortcut for the application in this directory, and you're done.

### Setting up the Application

When you first run the application, it will check for a file named `userconfig.xml` stored in your current user profile directory `(C:\Users\{username})`. If the file doesn't already exist, the application will create it and prompt you to enter your email ID and team name. Type the team name with caution as it is case sensitive and ensure you give your own email ID since you will be sent an OTP for email verification. This setup is a one-time activity. If your userconfig.xml file is created, you don't need to repeat the process. However, if you need to change your team name or email ID for any reason, you can remove the file and start from scratch.

### Clocking In and Out

The application allows users to clock in and clock out with minimal intervention.

#### Clocking In

When the application is running, it automatically prompts the user to clock in whenever Microsoft Teams is opened, providing yes or no choices. Upon clicking yes, the user is automatically clocked in. If the user clicks no, a popup will appear asking to set a reminder to clock in with default choices of 10, 15, and 30 minutes, and also allows the user to type any number in the textbox to set a custom reminder. If the application was started after already clocking in, it will track the user's time from that point without any hassle.

#### Clocking Out

When the application is running and Microsoft Teams is closed (process is killed), it prompts the user to clock out with yes or no options. Upon clicking yes, the user is automatically clocked out. If the user clicks no, the prompt to clock out will reappear after 10 seconds. For this functionality to work seamlessly, ensure you uncheck the "On close, keep the application running" setting in Teams.

## Summary

This application automates the process of clocking in and out, making it easier to manage your work hours with Microsoft Teams. By setting up the application once and letting it run in the background, users can focus on their work without worrying about manually tracking their time.
