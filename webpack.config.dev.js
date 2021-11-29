import webpack from 'webpack';
import path from "path";
const __dirname = path.resolve();

export default {
    target: "web",
    mode: "development",
    entry: ["babel-polyfill", "./src/main.js"],
    output: {
        path: path.resolve(__dirname, "dist"),
        filename: "JsToExcel.html.js",
        library: 'JsToExcel',
        libraryTarget: "window"
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