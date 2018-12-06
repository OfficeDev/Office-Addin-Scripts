import * as metadata from 'custom-functions-metadata';
import * as path from 'path';
import * as fs from 'fs';

const pluginName = 'customfunctions-plugin';

type CfType = {input:string, output:string};

class CustomFunctionsPlugin {
    options: CfType;
    constructor (options:CfType) {
        // Default options
        this.options = options;
    }

    //@ts-ignore
    apply (compiler) {

        const outputPath = compiler.options.output.path;

        if (compiler.hooks) {
            //@ts-ignore
            compiler.hooks.compile.tap(pluginName, (compilation) => {
                //Create dist folder if it doesn't exist
                try {
                    fs.mkdirSync(outputPath)
                    } 
                catch (err) {
                      if (err.code !== 'EEXIST') throw err
                    }
                const cfmetadata = metadata.generate(this.options.input, path.join(outputPath, this.options.output));
            });
        }
        else {
            console.log('hooks not found');
        }
    }

}

module.exports = CustomFunctionsPlugin;