# drumload
A simple VBScript to download a file, piecemeal if necessary.

## Usage
From the command prompt

    > cscript drumload.vbs http://www.example.com/percussion.zip drum.zip

You should see output along the lines of

> Microsoft (R) Windows Script Host Version 5.812  
> Copyright (C) Microsoft Corporation. All rights reserved.

> Downloading http://www.example.com/percussion.zip  
> Writing to C:\Users\blandbutgreasy\Desktop\drum.zip  
> File Size: 115683215 bytes  
> Downloaded: 13564243 / 115683215 bytes,  11.73%  
> Downloaded: 103378806 / 115683215 bytes,  89.36%  
> Downloaded: 115683215 / 115683215 bytes,  100.00%  
> Complete. File saved to C:\Users\blandbutgreasy\Desktop\drum.zip  
