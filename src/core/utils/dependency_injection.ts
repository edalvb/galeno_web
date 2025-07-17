import { Container } from "inversify";
import { Keys } from "./keys";
import {
  GeneralDataRepository,
  IGeneralDataRepository,
} from "../../app/shared/data/repositories/general_data_repository";

export const sl = new Container();

sl.bind<IGeneralDataRepository>(Keys.GENERAL_DATA_REPOSITORY)
  .to(GeneralDataRepository)
  .inSingletonScope();
