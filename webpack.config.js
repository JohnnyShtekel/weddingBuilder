var path = require('path');
var webpack = require('C:\\Users\\DEV\\AppData\\Roaming\\npm\\node_modules\\webpack');

var ROOT_PATH = path.resolve(__dirname, 'static');

module.exports = {
    entry: path.resolve(ROOT_PATH, 'src/js/index'),
    output: {
        path: path.join(ROOT_PATH, 'build'),
        filename: 'bundel.js'
    },

    module: {
        loaders: [
            {
                test: /\.css$/,
                exclude: /node_modules/,
                loaders: ['style', 'css'],
                include: path.resolve(ROOT_PATH, 'src/css')
            },
            {
                test: /\.(js|jsx|es6)$/,
                exclude: /node_modules/,
                loader: 'babel',


                query: {
                    presets: [
                        'babel-preset-es2015',
                        'babel-preset-react',
                        'babel-preset-stage-0',
                    ].map(require.resolve),


                }

            }
        ]
    }
};

