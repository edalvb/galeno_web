export class ResponseMutationModel<T> {
  success: boolean;
  message: string;
  data?: T;

  constructor(data: { success: boolean; message: string; data?: T }) {
    this.success = data.success;
    this.message = data.message;
    this.data = data.data;
  }
}
