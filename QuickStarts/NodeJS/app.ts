import GraphPagesAPI, { IPage } from './GraphController';
import config from './config';
import print from './logHelper';

const { siteId } = config;

// uncomment to activate each example:

// CopyPageToMultipleSites(/* page id */, [/* site ids */])
// DeletePageBeforeTargetDate(/* target date*/);
// PromotePagesAsNews([ /* page ids */]);


/**
 * Scenario #1: Copy a page to multiple sites
 *
 * @param {string} sourcePageId - source page id
 * @param {string[]} targetSiteIds - target site ids
 */
async function CopyPageToMultipleSites(sourcePageId: string, targetSiteIds: string[]) {
  const GraphPages = new GraphPagesAPI(config);
  const token = await GraphPages.getAuthenticationToken();
  print.logToken(token);
  GraphPages.storeToken(token);
 
  const sourcePage = await print.log("Get page content", GraphPages.getPage(siteId, sourcePageId));
  targetSiteIds
    .forEach(async targetSiteId => {
      const targetPage = modifyPage(sourcePage);
      const page = await print.log("Creating page", GraphPages.createPage(targetSiteId, targetPage));
      print.log("Publish page", GraphPages.publishPage(targetSiteId, page.id!));
    });
}

/**
 * Scenario #2: Delete pages before a target date
 * 
 * @param {Date} date - The target date
 */

async function DeletePageBeforeTargetDate(date: Date) {
  const GraphPages = new GraphPagesAPI(config);
  const token = await GraphPages.getAuthenticationToken();
  print.logToken(token);
  GraphPages.storeToken(token);

  const response = await print.log("List pages", GraphPages.listPages(siteId));
  const pageList: IPage[] = response.value;

  pageList
    .filter(page => new Date(page.lastModifiedDateTime!) < date)
    .forEach(async page => {
      print.log("Deleting page", GraphPages.deletePage(siteId, page.id!));
    });
}

/**
 * Scenario #3: Promote multiple pages as newposts
 *
 * @param {string[]} pageIds - Ids of the pages should be promoted
 */
async function PromotePagesAsNews(pageIds: string[]) {
  const GraphPages = new GraphPagesAPI(config);
  const token = await GraphPages.getAuthenticationToken();
  print.logToken(token);
  GraphPages.storeToken(token);

  pageIds
    .forEach(id => {
      print.log("Update page", GraphPages.updatePage(siteId, id, { promotionKind: 'newsPost' }));
    });
}

/**
 * Update this function to modify your own page
 * @param {IPage} sourcePage
 * @return {*}  {IPage}
 */
function modifyPage(sourcePage: IPage): IPage {
  return {
    ...sourcePage,
  };
}