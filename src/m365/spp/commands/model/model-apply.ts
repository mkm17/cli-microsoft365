import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spp } from '../../../../utils/spp.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
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
        option: '-u, --siteUrl <siteUrl>'
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

        return validation.isValidSharePointUrl(args.options.siteUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'title'] });
    this.optionSets.push({ options: ['listTitle', 'listId ', 'listUrl'] });
  }

  #initTypes(): void {
    this.types.string.push('siteUrl', 'id', 'title', 'listTitle', 'listId ', 'listUrl');
    this.types.boolean.push('defaultView');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.log(`Applying ${args.options.id || args.options.title} model to ${args.options.listId || args.options.listUrl || args.options.listTitle}...`);
      }

      const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);

      const requestOptions: CliRequestOptions = {
        url: `${siteUrl}/_api/machinelearning/publications`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        data: {
          Publications: {
            results: [
              {
                ModelUniqueId: args.options.id,
                TargetSiteUrl: args.options.title,
                TargetWebServerRelativeUrl: args.options.listId,
                TargetLibraryServerRelativeUrl: args.options.listTitle,
                ViewOption: args.options.listUrl
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

  private getCorrectRequestUrl(siteUrl: string, args: CommandArgs): string {
    if (args.options.id) {
      return `${siteUrl}/_api/machinelearning/models/getbyuniqueid('${args.options.id}')`;
    }

    return `${siteUrl}/_api/machinelearning/models/getbytitle('${formatting.encodeQueryParameter(args.options.title!)}')`;
  }
}

export default new SppModelRemoveCommand();