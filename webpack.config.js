"use strict";

module.exports = {
	mode: "development", //as opposed to production
	entry: "./src/app.js",
	output: {
		path: __dirname + "/builds/",
		filename: "bundle.js"
	},
	module: {
		rules: [
			{
				test: /\.js$/,
				exclude: /node_modules/,
				use: {
					loader: "babel-loader",
					options: {
						presets: [
							"env"
						]
					}
				}
			}
		]
	}
};
