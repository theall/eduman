set CUR_DIR=%CD%
set NODEMON_DIR=%CUR_DIR%\node_modules\.bin
set PATH=%NODEMON_DIR%;%PATH%
nodemon proxy.js