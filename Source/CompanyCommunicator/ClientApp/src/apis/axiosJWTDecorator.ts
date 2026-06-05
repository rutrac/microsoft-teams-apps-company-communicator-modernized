// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import axios, { AxiosResponse, AxiosRequestConfig } from "axios";
import { app, authentication } from "@microsoft/teams-js";
import i18n from '../i18n';

export class AxiosJWTDecorator {
    public async get<T = any, R = AxiosResponse<T>>(
        url: string,
        handleError: boolean = true,
        needAuthorizationHeader: boolean = true,
        config?: AxiosRequestConfig,
    ): Promise<R> {
        try {
            if (needAuthorizationHeader) {
                config = await this.setupAuthorizationHeader(config);
            }
            return await axios.get(url, config);
        } catch (error) {
            if (handleError) {
                this.handleError(error);
                throw error;
            }
            else {
                throw error;
            }
        }
    }

    public async delete<T = any, R = AxiosResponse<T>>(
        url: string,
        handleError: boolean = true,
        config?: AxiosRequestConfig
    ): Promise<R> {
        try {
            config = await this.setupAuthorizationHeader(config);
            return await axios.delete(url, config);
        } catch (error) {
            if (handleError) {
                this.handleError(error);
                throw error;
            }
            else {
                throw error;
            }
        }
    }

    public async post<T = any, R = AxiosResponse<T>>(
        url: string,
        data?: any,
        handleError: boolean = true,
        config?: AxiosRequestConfig
    ): Promise<R> {
        try {
            config = await this.setupAuthorizationHeader(config);
            return await axios.post(url, data, config);
        } catch (error) {
            if (handleError) {
                this.handleError(error);
                throw error;
            }
            else {
                throw error;
            }
        }
    }

    public async put<T = any, R = AxiosResponse<T>>(
        url: string,
        data?: any,
        handleError: boolean = true,
        config?: AxiosRequestConfig
    ): Promise<R> {
        try {
            config = await this.setupAuthorizationHeader(config);
            return await axios.put(url, data, config);
        } catch (error) {
            if (handleError) {
                this.handleError(error);
                throw error;
            }
            else {
                throw error;
            }
        }
    }

  private handleError(error: any): void {
    if (error.hasOwnProperty("response")) {
      const errorStatus = error.response.status;
      if (errorStatus === 403) {
        window.location.href = `/errorpage/403?locale=${i18n.language}`;
      } else if (errorStatus === 401) {
        window.location.href = `/errorpage/401?locale=${i18n.language}`;
      } else {
        window.location.href = `/errorpage?locale=${i18n.language}`;
      }
    } else {
      window.location.href = `/errorpage?locale=${i18n.language}`;
    }
  }

    private async setupAuthorizationHeader(
        config?: AxiosRequestConfig
    ): Promise<AxiosRequestConfig> {
        await app.initialize();
        try {
            const token = await authentication.getAuthToken();
            if (!config) {
                // axios.defaults is AxiosDefaults<any>; its `headers` field is HeadersDefaults
                // (with per-method slots .common/.get/.post/...). It is structurally compatible
                // with AxiosRequestConfig at runtime, but TS 5's stricter axios types reject
                // the direct assignment, hence the cast.
                config = axios.defaults as unknown as AxiosRequestConfig;
            }
            config.headers["Authorization"] = `Bearer ${token}`;
            config.headers["Accept-Language"] = i18n.language;
            return config;
        } catch (error) {
            // When getAuthToken fails, redirect to sign-in so the user can grant consent.
            console.error("Error from getAuthToken: ", error);
            window.location.href = `/signin?locale=${i18n.language}`;
            // Return a never-resolving promise so the outer try/catch in get/post/etc.
            // never fires handleError() (which would override this navigation with /errorpage).
            return new Promise<AxiosRequestConfig>(() => {});
        }
    }
}

const axiosJWTDecoratorInstance = new AxiosJWTDecorator();
export default axiosJWTDecoratorInstance;