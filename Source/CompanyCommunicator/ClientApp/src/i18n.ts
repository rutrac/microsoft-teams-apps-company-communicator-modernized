// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import i18n from 'i18next';
import { initReactI18next } from 'react-i18next';
import Backend from 'i18next-http-backend';
import dayjs from 'dayjs';
import localizedFormat from 'dayjs/plugin/localizedFormat';
import duration from 'dayjs/plugin/duration';
import utc from 'dayjs/plugin/utc';

dayjs.extend(localizedFormat);
dayjs.extend(duration);
dayjs.extend(utc);

export const defaultLocale = () => {
    return 'en-US';
} 

i18n
  // load translation using http -> see /public/locales (i.e. https://github.com/i18next/react-i18next/tree/master/example/react/public/locales)
  // learn more: https://github.com/i18next/i18next-http-backend
  .use(Backend)
  // pass the i18n instance to react-i18next.
  .use(initReactI18next)
  // init i18next
  // for all options read: https://www.i18next.com/overview/configuration-options
    .init({
    fallbackLng: defaultLocale(),
    interpolation: {
      escapeValue: false, // not needed for react as it escapes by default
    }
  });

// Map i18next locale codes (used in URL ?locale=) to dayjs locale identifiers.
// dayjs locales are loaded on demand via dynamic import so we don't bundle all ~140 locales eagerly.
const dayjsLocaleMap: Record<string, string> = {
    'ar-SA': 'ar-sa',
    'de-DE': 'de',
    'en-US': 'en',
    'es-ES': 'es',
    'fr-FR': 'fr',
    'he-IL': 'he',
    'it-IT': 'it',
    'ja-JP': 'ja',
    'ko-KR': 'ko',
    'pt-BR': 'pt-br',
    'pt-PT': 'pt',
    'ru-RU': 'ru',
    'zh-CN': 'zh-cn',
    'zh-TW': 'zh-tw',
};

const loadDayjsLocale = async (locale: string): Promise<string> => {
    const dayjsCode = dayjsLocaleMap[locale] ?? 'en';
    if (dayjsCode === 'en') {
        dayjs.locale('en');
        return 'en';
    }
    try {
        await import(`dayjs/locale/${dayjsCode}.js`);
        dayjs.locale(dayjsCode);
        return dayjsCode;
    } catch {
        dayjs.locale('en');
        return 'en';
    }
};

export const updateLocale = () => {
    const search = window.location.search;
    const params = new URLSearchParams(search);
    const locale = params.get("locale") || defaultLocale();
    i18n.changeLanguage(locale);
    void loadDayjsLocale(locale);
};

export const formatDate = (date: string) => {
    if (!date) return date;
    return dayjs(date).format('l LT');
}

export const formatDuration = (startDate: string, endDate: string) => {
    let result = "";
    const search = window.location.search;
    const params = new URLSearchParams(search);
    const locale = params.get("locale") || defaultLocale();
    if (startDate && endDate) {
        const difference = dayjs(endDate).diff(dayjs(startDate));
        const totalDuration = dayjs.duration(difference);
        // Handling the scenario of duration being more than 24 hrs as it is not done natively.
        const hh = ("0" + Math.floor(totalDuration.asHours())).slice(-2);
        const dayjsCode = dayjsLocaleMap[locale] ?? 'en';
        result = formatNumber(parseInt(hh)) + dayjs.utc(totalDuration.asMilliseconds()).locale(dayjsCode).format(":mm:ss")
    }
    return result;
}

export const formatNumber = (number: any) => {
    const search = window.location.search;
    const params = new URLSearchParams(search);
    const locale = params.get("locale") || defaultLocale();
    return Number(number).toLocaleString(locale);
}

export default i18n;