/**
 * Method to reduce repetitive requests to the Graph utilized in Batch
 *
 * @class BatchRequest
 */
export class BatchRequest {
  /**
   * url used in request
   *
   * @type {string}
   * @memberof BatchRequest
   */
  public resource: string;
  /**
   * method passed to be requested
   *
   * @type {string}
   * @memberof BatchRequest
   */
  public method: string;
  public constructor(resource: string, method: string) {
    if (resource.charAt(0) !== '/') {
      resource = '/' + resource;
    }
    this.resource = resource;
    this.method = method;
  }
}
