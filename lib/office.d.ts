export {};

declare global {
  interface Window {
    Office?: typeof Office;
  }

  namespace Office {
    const context: {
      mailbox: {
        item?: MailboxItem;
      };
    };

    function onReady(
      callback: (info: { host: string; platform: string }) => void,
    ): Promise<unknown>;

    const AsyncResultStatus: { Succeeded: string; Failed: string };

    namespace MailboxEnums {
      const ComposeType: {
        NewMail: string;
        Reply: string;
        Forward: string;
      };
    }

    namespace actions {
      function associate(
        name: string,
        handler: (event: OnSendEvent) => void,
      ): void;
    }

    interface AsyncResult<T> {
      status: string;
      value: T;
      error?: { message: string };
    }

    interface CustomProperties {
      get(name: string): string | undefined;
      set(name: string, value: string): void;
      saveAsync(callback?: (result: AsyncResult<void>) => void): void;
    }

    interface CoercionType {
      Html: string;
      Text: string;
    }

    interface MailboxItem {
      subject: {
        getAsync(callback: (result: AsyncResult<string>) => void): void;
      };
      body: {
        getAsync(
          coercionType: string,
          callback: (result: AsyncResult<string>) => void,
        ): void;
        setAsync(
          data: string,
          options: { coercionType: string },
          callback?: (result: AsyncResult<void>) => void,
        ): void;
        prependAsync(
          data: string,
          options: { coercionType: string },
          callback?: (result: AsyncResult<void>) => void,
        ): void;
      };
      getComposeTypeAsync(
        callback: (result: AsyncResult<{ composeType: string }>) => void,
      ): void;
      loadCustomPropertiesAsync(
        callback: (result: AsyncResult<CustomProperties>) => void,
      ): void;
      internetHeaders: {
        setAsync(
          headers: Record<string, string>,
          callback?: (result: AsyncResult<void>) => void,
        ): void;
      };
    }

    interface OnSendEvent {
      completed(options: {
        allowEvent: boolean;
        errorMessage?: string;
        cancelLabel?: string;
        commandId?: string;
      }): void;
    }

    const CoercionType: CoercionType;
  }
}
