import { Consts } from '../consts/Consts';
import { PNPService } from './PNPService';
import { Text } from '@microsoft/sp-core-library';
import { UrlHelper } from '../helpers/UrlHelper';
export  interface ITerm {
  id : string;
    labels: {
        name: string;
        languageTag: string;
    }[];
  }
export class TaxonomyService {
  constructor() {
    //empty
  }
  private  terms: ITerm[] = [];

  public  async init(pnpService:PNPService) {
    this.terms = await this.retrieveTerms(pnpService);
  }

  private  async retrieveTerms(pnpService:PNPService) {

    const taxonomyListUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(),
      Consts.LISTS.TAXONOMY_HIDDEN);

    const taxonomyResponse = await pnpService.getListByUrl(taxonomyListUrl)
      .items
      .select(
        Consts.FIELDS.TAXONOMY.TERM_ID,
        Consts.FIELDS.TAXONOMY.TERM_LABEL_EN,
        Consts.FIELDS.TAXONOMY.TERM_LABEL_FR)
      .top(1000)();

    return taxonomyResponse.map((term): ITerm => ({
      id: term[Consts.FIELDS.TAXONOMY.TERM_ID],
      labels: [
        {
          name: term[Consts.FIELDS.TAXONOMY.TERM_LABEL_EN],
          languageTag: 'en-us'
        },
        {
          name: term[Consts.FIELDS.TAXONOMY.TERM_LABEL_FR],
          languageTag: 'fr-fr'
        }
      ]
    }));
  }

  public getLocalizedTerm(termId:string, language: string) {
    let fetchedTerm: ITerm = null;
    for (let i = 0; fetchedTerm === null && i < this.terms.length; i++) {
      fetchedTerm = this.fetchTerm(this.terms[i], termId);
    }
    return fetchedTerm?.labels.find((t) => t.languageTag.toLowerCase() === language).name;
  }

  private fetchTerm(currentTerm: ITerm, termId:string):ITerm {
    if (termId  === currentTerm.id || termId === Text.format('0{0}', currentTerm.id)) return currentTerm;
    return null;
  }
}
