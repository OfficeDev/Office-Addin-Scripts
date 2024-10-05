/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file is the main Node.js server file that defines the express middleware.
 */

if (process.env.NODE_ENV !== "production") {
  require("dotenv").config();
}
import * as createError from "http-errors";
import express from "express";
import * as path from "path";
import * as cookieParser from "cookie-parser";
import logger from "morgan";
import { getGraphData } from "./msgraph-helper";
import { getAccessToken, validateJwt } from "./ssoauth-helper";
import { usageDataObject } from "../defaults";

/* global process, require, __dirname */

export class App {
  appInstance: express.Express;
  port: number | string;
  constructor(port: number | string) {
    this.appInstance = express();
    this.port = port;
  }

  async initialize() {
    this.appInstance.set("port", this.port);

    // // view engine setup
    this.appInstance.set("views", path.join(__dirname, "views"));
    this.appInstance.set("view engine", "pug");

    this.appInstance.use(logger("dev"));
    this.appInstance.use(express.json());
    this.appInstance.use(express.urlencoded({ extended: false }));
    this.appInstance.use(cookieParser());

    /* Turn off caching when developing */
    if (process.env.NODE_ENV !== "production") {
      this.appInstance.use(express.static(path.join(process.cwd(), "dist"), { etag: false }));

      this.appInstance.use(function (req, res, next) {
        res.header("Cache-Control", "private, no-cache, no-store, must-revalidate");
        res.header("Expires", "-1");
        res.header("Pragma", "no-cache");
        next();
      });
    } else {
      // In production mode, let static files be cached.
      this.appInstance.use(express.static(path.join(process.cwd(), "dist")));
    }

    const indexRouter = express.Router();
    indexRouter.get("/", function (req, res) {
      res.render("/taskpane.html");
    });

    this.appInstance.use("/", indexRouter);

    // listen for 'ping'
    this.appInstance.get("/ping", function (req: any, res: any) {
      res.send(process.platform);
    });

    this.appInstance.get("/taskpane.html", async (req: any, res: any) => {
      return res.sendfile("taskpane.html");
    });

    this.appInstance.get("/fallbackauthdialog.html", async (req: any, res: any) => {
      return res.sendfile("fallbackauthdialog.html");
    });

    this.appInstance.get("/getuserdata", validateJwt, async function (req: any, res: any, next: any) {
      const graphTokenResponse = await getAccessToken(req.get("Authorization"));
      if (graphTokenResponse.claims || graphTokenResponse.error) {
        graphTokenResponse.claims
          ? usageDataObject.reportSuccess("getuserdata()", { details: "Got claims response" })
          : usageDataObject.reportError(
              "AccessTokenError",
              new Error("Access Token Error: " + graphTokenResponse.error)
            );
        res.send(graphTokenResponse);
      } else {
        const graphToken: string = graphTokenResponse.access_token;
        const graphUrlSegment: string = process.env.GRAPH_URL_SEGMENT || "/me";
        const graphQueryParamSegment: string = process.env.QUERY_PARAM_SEGMENT || "";

        const graphData = await getGraphData(graphToken, graphUrlSegment, graphQueryParamSegment);

        // If Microsoft Graph returns an error, such as invalid or expired token,
        // there will be a code property in the returned object set to a HTTP status (e.g. 401).
        // Relay it to the client. It will caught in the fail callback of `makeGraphApiCall`.
        if (graphData.code) {
          usageDataObject.reportException("getuserdata()", graphData.code);
          next(createError(graphData.code, "Microsoft Graph error " + JSON.stringify(graphData)));
        } else {
          usageDataObject.reportSuccess("getuserdata()");
          res.send(graphData);
        }
      }
    });

    // Catch 404 and forward to error handler
    this.appInstance.use(function (req: any, res: any, next: any) {
      next(createError(404));
    });

    // error handler
    this.appInstance.use(function (err: any, req: any, res: any) {
      // set locals, only providing error in development
      res.locals.message = err.message;
      res.locals.error = req.app.get("env") === "development" ? err : {};

      // render the error page
      res.status(err.status || 500);
      res.render("error");
    });
  }
}
