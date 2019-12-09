## Document Management UI

An SPFx webpart designed to monitor and interact with the documents related to the IEP/504 Educational Plan Documentation Solution.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options
* ```gulp clean ```                    Removes artifacts which may impact proper runtime.
* ```gulp build ```                    Builds the runtime and performs code analysis.
* ```gulp serve --nobrowser ```        Starts the local debug server.
* ```gulp bundle --ship ```            Compresses and minifies code; use prior to packaging.
* ```gulp package-solution --ship ```  Creates an .sppkg file for distribution to a SharePoint App Catalog.




### Debug
gulp serve --nobrowser
