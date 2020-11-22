
## environment setup
###  nodejs
need nvm to mangeme the nodejs envorinment, the react app resides in front directory.
NodeJS version 12 or above

### python
Python version 3.8 or above
use the python3 venv or conda to manage the virtual environment.
package was recored in the requirements.txt, use pip freeze > requirements to update this file

### IDE
vscode is recommended.
PyCharm and WebStorm is also ok.


### front end project
created with create-react-app
use the reduxjs to management the global state.
use the react-route to handle the page routing.

report config was request from the backend service, default url was configured in the feature/reports/reportsSlice.js line 7 config_url

### backend project
defautl run with flask run, the app config was in the config directory.
