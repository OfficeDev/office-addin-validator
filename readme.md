# Microsoft Office Manifest Validator

## How to install
```bash
npm install -g manifest-validator@0.0.1-beta.2
```

## How to run
```bash
$ validate-office-addin <your_manifest.xml>
```

### Command Line Options
```bash
-l or --language
```
Allows you to localize the response.
* Type: String
* Default: en-US
* Optional

```bash
-h or --help
```
Allows you to see all options
* Type: Boolean
* Default: False
* Optional

## How to develop
1. Open bash terminal or shell
2. Navigate to the folder you want to install this tool
3. Run the following commands:

```bash
$ git clone https://github.com/OfficeDev/manifest-validator
$ cd manifest-validator
$ npm install
$ npm start
```

4. Open a new tab in your terminal. Make sure you are still in your project directory.
5. Run the following command:

```bash
$ npm link
```
