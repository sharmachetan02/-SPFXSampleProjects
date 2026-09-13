
class DomHelper {

  public static getWindowWidth(): number {
    const docElemProp = window.document.documentElement.clientWidth,
      body = window.document.body;
    return window.document.compatMode === 'CSS1Compat' && docElemProp || body && body.clientWidth || docElemProp;

  }
  public static getPosition(): Promise<GeolocationPosition> {
    const options = {
      enableHighAccuracy: false,
      timeout: 5000,
      maximumAge: 1000 * 60 * 60
    };
    return new Promise((resolve, reject) =>
      navigator.geolocation.getCurrentPosition(resolve, reject, options)
    );
  }
  public static async delay(ms) {
    return new Promise((resolve) => {
      setTimeout(resolve, ms);
    });
  }
  public static isEllipsisActive(element) {
    if (!element) return false;
    return (element.offsetWidth < element.scrollWidth);
  }
  public static openUrl(url: string) {
    const navState = { url };
    history.pushState(navState, null, url);
    const newPopState = new PopStateEvent('popstate', { state: navState });
    window.dispatchEvent(newPopState);
  }
  public static cleanRichHtml(inputHtml: string): string {
    if (!inputHtml) return '';

    // Parse input HTML in a safe DOM
    const doc = new DOMParser().parseFromString(inputHtml, 'text/html');

    const walk = (node: Node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        // Normalize non-breaking spaces
        const t = node as Text;
        t.data = t.data.replace(/\u00A0/g, ' ');
        return;
      }

      if (node.nodeType === Node.ELEMENT_NODE) {
        const el = node as HTMLElement;

        // Remove empty block elements
        if (['P', 'DIV', 'SPAN'].includes(el.tagName)) {
          const text = el.textContent?.replace(/\u00A0/g, ' ').trim() ?? '';
          if (!text) {
            el.remove();
            return; // stop walking removed node
          }
        }

        // Collapse consecutive <br> into a single
        Array.from(el.childNodes).forEach((child) => {
          if (child.nodeName === 'BR' && child.previousSibling?.nodeName === 'BR') {
            child.remove();
          }
        });

        // Recurse children
        Array.from(el.childNodes).forEach(walk);
      }
    };

    Array.from(doc.body.childNodes).forEach(walk);

    // Remove trailing <br> at the end
    while (doc.body.lastChild && doc.body.lastChild.nodeName === 'BR') {
      doc.body.lastChild.remove();
    }

    return doc.body.innerHTML.trim();
  }
}
export default DomHelper;

