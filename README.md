1. Download a chrome driver which version is same as your chrome version.
   View your chrome version : chrome://settings/help
   Then download chrome driver to somewhere eg : C:\Program Files\Google\Chrome\Application\

2. Open cmd 
   cd C:\Program Files\Google\Chrome\Application
   "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:/path/to/your/profile"
   It will prompt a chrome browser with port 9222

3. Then close other chrome browser in your desktop. Input Remedy url in the port 9222 chrome browser.
