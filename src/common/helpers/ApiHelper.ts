import * as xml2js from 'xml2js';
export enum CallMethod{
  GET,
  POST
}
export class ApiHelper {
  public static callMethod = CallMethod;
  public static returnedTypes = {
    json: 'application/json; odata=verbose',
    xml: 'text/xml; odata=verbose'
  };
  public static referrerPolicies = {
    noReferrer: 'no-referrer-when-downgrade' as ReferrerPolicy,
    strictOrigin: 'strict-origin-when-cross-origin' as ReferrerPolicy,
  };
  public static async callApi(endpoint:string, request:RequestInit) {
    const promise = await new Promise((resolve, reject) => {

      fetch(endpoint, request).then((response: Response) => {
        if (response.ok) {
          response.clone().json().then((data: unknown) => {
            resolve(data);
          }).catch(() => {

            response.text().then((data: string) => {
              ApiHelper.parseXml(data).then((parsedXml: unknown) => {
                resolve(parsedXml);
              }).catch((error) => { reject(error); });
            }).catch((error) => { reject(error); });
          });
        } else {
          reject(response);
        }
      });

    });
    return promise;

  }

  public static async parseXml(xml:string) {
    const promise = await new Promise((resolve, reject) => {
      const parser = new xml2js.Parser({ explicitArray: false });
      parser.parseString(xml, (error, result) => {
        if (error) reject(error);
        else resolve(result);
      });
    });
    return promise;
  }
}
