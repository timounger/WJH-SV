:: "packages.bat"
:: install packages for youtube downloader

:: -r, --requirement <file>    Install from the given requirements file.
:: -c, --constraint <file>     Constrain versions using the given constraints file
:: -U, --upgrade               Upgrade all specified packages to the newest available version
:: --no-cache-dir              Disable the cache.
:: --use-deprecated            Enable deprecated functionality, that will be removed in the future.

:: pip freeze > requirements.txt
:: pip uninstall -r requirements.txt -y

python.exe -m pip install --upgrade pip

python.exe -m pip install --upgrade --no-cache-dir --use-deprecated=legacy-resolver --requirement packages.txt

pause