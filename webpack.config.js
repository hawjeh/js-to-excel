import webpack from 'webpack';
import path from "path";
const __dirname = path.resolve();

export default {
    mode: "production",
    entry: ["babel-polyfill", "./src/main.js"],
    output: {
        filename: "JsToExcel.min.js",
        path: path.resolve(__dirname, "dist"),
        library: 'JsToExcel',
        libraryTarget: "var"
    },
    plugins: [
        new webpack.optimize.LimitChunkCountPlugin({
            maxChunks: 1,
        }),
    ],
    module: {
        rules: [
            {
                test: /\.js$/,
                exclude: [
                    /(node_modules)/,
                    path.resolve(__dirname, './src/test/')
                ],
                use: {
                    loader: "babel-loader",
                    options: {
                        presets: ["@babel/preset-env"]
                    }
                }
            }
        ]
    },
    experiments: {
        outputModule: true,
    },
};