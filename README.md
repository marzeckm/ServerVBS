# ServerVBS
ServerVBS is a static webserver program that is executed by a VBScript file. It uses TinyWeb by Maxim Masiutin as its kernel.

## Installation
To install ServerVBS, simply download the files from GitHub and run Server.vbs in the main folder. The program will ask if you want to download and install the missing files (`TinyWeb.exe`, `TinySSL.exe`, `Libeay32.dll`, and `Libssl32.dll`) to the `/bin` folder. If you confirm, the files will be downloaded and installed automatically. Otherwise, you can manually install the files to the bin folder in the main folder.

You can add HTTPS support by adding a `key.pem` and a `cert.pem` file to the bin folder. These files can be created with `OpenSSL`. To enable SSL, change the ssl setting in the `/bin/config.inf` file from `ssl:no` to `ssl:yes`.

## Configuration
You can configure ServerVBS by changing the settings in the `/bin/config.inf` file, located in the bin folder. The available settings are:

`port`: the port number to listen on (default is `80`).  
`ssl`: whether to enable SSL (default is `no`).  
`language`: the language to use for error messages (default is English (`en`), but German (`de`) is also available).  
`homepath`: the root directory of your website (default is `./www`).  

## Usage
When the server is started, you can use the following commands:

`start`: start the server after it has been paused.  
`pause`: pause the server without closing the program.  
`stop`: stop the server and close the program.  
`restart`: restart the server.  

## Credits
This program is based on the work of Maxim Masiutin, who developed TinyWeb. (http://www.ritlabs.com/en/products/tinyweb/)  
`Libeay32.dll` and `Libssl32.dll` from OpenSSL, distributed under the Apache License 2.0.

## Requirements
- Windows 7, Windows Server 2008 R2, Windows 8, Windows 8.1, Windows Server 2012, Windows 10, Windows Server 2016, Windows Server 2019, Windows 11, Windows Server 2022  
- The execustion of VBScript must be activated. Many Antiviruses block the execution of VBScript.

## Contribute
If you want to contribute to the development of this project, feel free to submit pull requests or open issues. Let's make ServerVBS even better together!

## License
This project is licensed under the [MIT License](LICENSE).
