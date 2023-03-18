# Excel-VBA
This repository contains a set of VBA modules that can be used in Excel workbooks.
- Modules are intended to be able to be used together.

## Functions
- TravelDistance(origin, destination)
  - This returns the distance in miles between the `origin` and `destination`.
- TravelTime(origin, destination)
  - This returns the travel time, in minutes, between the `origin` and `destination`.
  - Travel time may vary based on current travel conditions.

## Bing Map API Key
The module assumes that you have stored your Bing Map API key is an environment variable named `BingMapAPiKey`.

## Performance Improvement
Performance can be slow if there are multiple distance of time references is an Excel workbook.
- Recommendation is to set calculation mode to manual.
- Calculation will occur on a save.
