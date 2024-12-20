# Android Platform Tools Updater

> [!CAUTION]
> Don't run random scripts from the internet.
>
> * This script is provided as-is, you should run it with caution.
> * If it breaks anything on your computer it's your fault.

<!-- Make Markdownlint not complain about MD028 -->

> [!NOTE]
> I recommend using [Scoop](https://scoop.sh/#/apps?q=adb+fastboot) or [WinGet](https://github.com/microsoft/winget-pkgs/tree/master/manifests/g/Google/PlatformTools) instead.

## Disclaimer

This is just a hobby project of mine. I will not update it regularely, and I can't promise fixes within reasonable time if anything breaks.

* You are youself responsible for what you do on your computer.
* I do not provide guarantees or support if something goes wrong or stops working.

That being said, I'm happy for any feedback you might have.

## Introduction

This little script fetches the latest version of "platform-tools" (ADB and Fastboot) by Google for Windows, and installs it either in system or user context.

It can be rerun later, if a new version is available from Google, it will update installed platform-tools.

## Requirements

The script is made to run on Windows 10 from PowerShell ISE using PowerShell 5.1.
