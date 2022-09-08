class CustomElementHelper {
  private readonly defaultPrefix = 'mgt';
  private _disambiguation = '';
  public withDisambiguation(disambiguation: string) {
    this._disambiguation = disambiguation;
    return this;
  }
  public get prefix(): string {
    return this._disambiguation ? `${this.defaultPrefix}-${this._disambiguation}` : this.defaultPrefix;
  }
  public get isDisambiguated(): boolean {
    return Boolean(this._disambiguation);
  }
}

const customElementHelper = new CustomElementHelper();
export { customElementHelper };
