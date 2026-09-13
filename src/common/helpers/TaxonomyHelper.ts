import { ITermInfo } from "@pnp/sp/taxonomy";
import { PNPService } from "../services/PNPService";

export class TaxonomyHelper {

  private static buildTreeData(terms: ITermInfo[], parentId: string): ITermInfo[] {
    const children = terms.filter((term) => {
      const termParentId = term.localProperties[0]['properties'].find((p) => p.key === 'ParentId')?.value;
      return termParentId === parentId;
    });
    const parent = terms.find((t) => t.id === parentId);
    const order = parent?.customSortOrder[0]?.order;
    const ordredChildren = order ? order.map((id) => children.find((t) => t.id === id)) : children;
    return ordredChildren.map((child) => {
      if (!child) return;
      console.log('child', child);
      return {
        ...child,
        children: this.buildTreeData(terms, child.id)
      };
    }
    );
  }

  public static async getTermSetTerms(pnpService: PNPService, termsetId: string): Promise<ITermInfo[]> {
    const terms = await pnpService.getTermsFromTermSet(termsetId)
      .select('id, labels, localProperties, CustomSortOrder')();
    return this.buildTreeData(terms, termsetId);
  }
  public static async getSubTerms(pnpService: PNPService, termsetId: string, termId: string): Promise<ITermInfo[]> {
    const terms = await pnpService.getTermsFromTermSet(termsetId)
      .select('id, labels, localProperties, CustomSortOrder')();
    return this.buildTreeData(terms, termId);

  }
}
