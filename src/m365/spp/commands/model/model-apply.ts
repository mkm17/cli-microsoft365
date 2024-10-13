import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spp, SppModel } from '../../../../utils/spp.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { ListInstance } from '../../../spo/commands/list/ListInstance.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  id?: string;
  title?: string;
  listTitle?: string;
  listId?: string;
  listUrl?: string;
  defaultView?: boolean;
}

class SppModelRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.MODEL_APPLY;
  }

  public get description(): string {
    return 'Applies (or syncs) a trained document understanding model to a document library';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        defaultView: !!args.options.defaultView
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--siteUrl <siteUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '--listTitle [listTitle]'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--defaultView'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.listId &&
          !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.siteUrl) && validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'title'] });
    this.optionSets.push({ options: ['listTitle', 'listId ', 'listUrl'] });
  }

  #initTypes(): void {
    this.types.string.push('siteUrl', 'webUrl', 'id', 'title', 'listTitle', 'listId ', 'listUrl');
    this.types.boolean.push('defaultView');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.log(`Applying ${args.options.id || args.options.title} model to ${args.options.listId || args.options.listUrl || args.options.listTitle}...`);
      }

      const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
      await spp.assertSiteIsContentCenter(siteUrl);

      const model = await this.getModel(siteUrl, args);
      const listInstance = await this.getListInfo(args);

      const requestOptions: CliRequestOptions = {
        url: `${siteUrl}/_api/machinelearning/publications`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        data: {
          Publications: {
            results: [
              {
                ModelUniqueId: model.UniqueId,
                TargetSiteUrl: siteUrl,
                TargetWebServerRelativeUrl: urlUtil.getServerRelativePath(args.options.webUrl, ''),
                TargetLibraryServerRelativeUrl: urlUtil.getServerRelativePath(args.options.webUrl, listInstance.Url),
                ViewOption: args.options.defaultView ? "NewViewAsDefault" : undefined
              }
            ]
          }
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getModel(siteUrl: string, args: CommandArgs): Promise<SppModel> {
    const requestOptions: CliRequestOptions = {
      url: this.getCorrectRequestUrl(siteUrl, args),
      headers: {
        accept: 'application/json;odata=nometadata'
      }
    };

    return request.get(requestOptions);
  }

  private getCorrectRequestUrl(siteUrl: string, args: CommandArgs): string {
    if (args.options.id) {
      return `${siteUrl}/_api/machinelearning/models/getbyuniqueid('${args.options.id}')`;
    }

    return `${siteUrl}/_api/machinelearning/models/getbytitle('${formatting.encodeQueryParameter(args.options.title!)}')`;
  }

  private getListInfo(args: CommandArgs): Promise<ListInstance> {
    let requestUrl = `${args.options.webUrl}/_api/web`;

    if (args.options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get<ListInstance>(requestOptions);
  }
}

export default new SppModelRemoveCommand();