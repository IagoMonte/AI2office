/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/";

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const httpsOptions = await devCerts.getHttpsServerOptions();
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.js",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
          {
            from: "node_modules/pdfjs-dist/build/pdf.worker.min.mjs",
            to: "pdf.worker.min.mjs",
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: httpsOptions,
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      setupMiddlewares: (middlewares, devServer) => {
        const http = require("http");
        const https = require("https");
        const url = require("url");

        devServer.app.all("/api-proxy", (req, res) => {
          const parsedReq = url.parse(req.url, true);
          const targetUrl = parsedReq.query.target;
          if (!targetUrl) {
            res.statusCode = 400;
            res.end(JSON.stringify({ error: "Missing target" }));
            return;
          }

          const parsed = url.parse(targetUrl);
          const isHttps = parsed.protocol === "https:";
          const lib = isHttps ? https : http;

          const proxyHeaders = Object.assign({}, req.headers);
          delete proxyHeaders.host;
          delete proxyHeaders.connection;
          proxyHeaders.host = parsed.host;

          const proxyReq = lib.request(
            {
              hostname: parsed.hostname,
              port: parsed.port || (isHttps ? 443 : 80),
              path: parsed.path,
              method: req.method,
              headers: proxyHeaders,
            },
            (proxyRes) => {
              res.writeHead(proxyRes.statusCode, proxyRes.headers);
              proxyRes.pipe(res);
            }
          );

          proxyReq.on("error", (err) => {
            res.statusCode = 502;
            res.end(JSON.stringify({ error: err.message }));
          });

          req.pipe(proxyReq);
        });

        return middlewares;
      },
    },
  };

  return config;
};
