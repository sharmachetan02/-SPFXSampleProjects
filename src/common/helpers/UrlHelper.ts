import { Consts } from '../consts/Consts';
import { Text } from '@microsoft/sp-core-library';

export class UrlHelper {

  public static getQueryStringParam(field: string, url: string) {
    const href = url ? url : window.location.href;
    const reg = new RegExp('[?&#]' + field + '=([^&#]*)', 'i');
    const qs = reg.exec(href);
    return qs ? qs[1] : null;
  }

  public static removeQueryStringParam(field: string, sourceURL: string) {
    let rtn = sourceURL.split('?')[0];
    let param = null;
    let paramsArr = [];
    const queryString =
      sourceURL.indexOf('?') !== -1 ? sourceURL.split('?')[1] : '';

    if (queryString !== '') {
      paramsArr = queryString.split('&');
      for (let i = paramsArr.length - 1; i >= 0; i -= 1) {
        param = paramsArr[i].split('=')[0];
        if (param === field) {
          paramsArr.splice(i, 1);
        }
      }

      if (paramsArr.length > 0) {
        rtn = rtn + '?' + paramsArr.join('&');
      }
    }
    return rtn;
  }

  public static addOrReplaceQueryStringParam(
    url: string,
    param: string,
    value: string
  ) {
    const re = new RegExp('[\\?&]' + param + '=([^&#]*)');
    const match = re.exec(url);
    let delimiter;
    let newString;

    if (match === null) {
      const hasQuestionMark = /\?/.test(url);
      delimiter = hasQuestionMark ? '&' : '?';
      newString = url + delimiter + param + '=' + value;
    } else {
      delimiter = match[0].charAt(0);
      newString = url.replace(re, delimiter + param + '=' + value);
    }

    return newString;
  }

  public static getListWebRelativeUrl(webAbsoluteUrl: string, listUrl: string) {
    const  tenantUrl = window.location.origin.toLowerCase();
    const relativeSiteUrl = webAbsoluteUrl.toLowerCase().replace(tenantUrl, '');
    return `${relativeSiteUrl}/${listUrl}`;
  }

  public static getDelvePictureUrl(email: string, size:string) {
    const emailEncoded = encodeURIComponent(email);
    return Text.format(Consts.OFFICE.OFFICE_DELVE_PROFILE_IMAGE_BASE_URL, size, emailEncoded);
  }

  public static navigate(link: string, fullPageReload = false) {
    //Including the protocol that you're using
    // if (link.search(/^http[s]?\\:\/\//) === -1) {
    //   //Not starting with a /
    //   if (link.search(/^\//) === -1) {
    //     link = '/' + link;
    //   }
    // }


    const isLayoutPage = location.href.toLowerCase().indexOf('/_layouts/') !== -1;
    // Check if we can bind into the SPFx navigation APIs
    if (!fullPageReload && history && (window['PopStateEvent'] || window['Event']) && window['dispatchEvent']) {
      // Create the new navigation state
      const navState = { url: link };

      // Adds the new navigation state to the browser history
      history.pushState(navState, null, link);

      // Check to trigger SharePoint navigation handler to partially reload the page
      let newPopState = null;
      const popStateString = 'popstate';
      const eventString = 'Event';
      const stateString = 'state';
      try {
        if (window['PopStateEvent']) {
          newPopState = new PopStateEvent(popStateString, { state: navState });
        }

        if (window[eventString] && !newPopState) {
          newPopState = new Event(popStateString);
          newPopState[stateString] = navState;
        }

        if (!newPopState) {
          newPopState = new PopStateEvent(popStateString, {
            bubbles: false,
            cancelable: true,
            state: navState
          });
        }
      } catch (e) {
        newPopState = document.createEvent(eventString);
        newPopState.initEvent(popStateString, false, true);
        newPopState[stateString] = navState;
      }

      if (newPopState) {
        const isDispatched = window.dispatchEvent(newPopState);
        if (isDispatched) {
          // Check if browser was loaded on the layouts page, if that was the case, trigger history go API
          if (isLayoutPage) {
            history.go();
          }
          return;
        }
      }
    }

    // Worst case, redirect the old way
    //location.href = link;
    window.open(link, '_blank');

  }

  public static onNavigationChange(callback: (path?: string) => void): void {
    ((history) => {
      const pushState = history.pushState;
      history.pushState = (state, key, path) => {
        pushState.apply(history, [state, key, path]);
        callback(typeof path === 'string' ? path : path?.toString());
      };
    })(window.history);

    window.addEventListener('popstate', () => {
      callback();
    });
  }
}

export enum PageOpenBehavior {
  'Self',
  'NewTab'
}
