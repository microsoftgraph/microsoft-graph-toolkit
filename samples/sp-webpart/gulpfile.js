'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

build.configureWebpack.mergeConfig({
    additionalConfiguration: (generatedConfiguration) => {
      generatedConfiguration.module.rules.push(
        {
          test: /\.m?js$/, use:
          {
            loader: "babel-loader",
            options:
            {
              presets: [["@babel/preset-env",
                {
                  targets: {
                    "ie": "11"
                  }
                }]]
            }
          }
        }
      );
  
      return generatedConfiguration;
    }
  });

build.initialize(gulp);
