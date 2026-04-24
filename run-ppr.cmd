@echo off
set AR_EMAIL=
set AR_PASSWORD=
set AR_CHROME_PROFILE=C:\PPRChrome
set AR_SHOW_UI=1
set AR_APPS_SCRIPT_URL=https://script.google.com/macros/s/AKfycbzkDWQojxVLChOzHjeiQajIKcKayD79-t1pWd4kt5KzIO_MN-G4BsBvRfBMMWkQ25Af/exec
rem If you captured a version in DevTools, uncomment and paste it:
rem set AR_VERSION=2025/11/07 08:23:20 PST

set CHROME_FLAGS=--disable-features=BlockThirdPartyCookies,ThirdPartyStoragePartitioning,PartitionConnectionsByNetworkIsolationKey,PartitionDomainReliabilityByNetworkIsolationKey,CrossSiteDocumentBlockingAlways --disable-features=SameSiteByDefaultCookies,CookiesWithoutSameSiteMustBeSecure

set AR_DEBUG=1
set AR_SNAP_MODE=none

node "%~dp0ppr.js"
pause
