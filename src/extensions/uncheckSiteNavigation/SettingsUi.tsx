import * as React from 'react';
import * as ReactDom from 'react-dom';
import { SettingsBar } from './SettingsBar';
import { SettingsPanel } from './SettingsPanel';

export interface ISettingsUiProps {
  enabled: boolean;
  canEdit: boolean;
  /** Open the panel immediately (deep link via query string). */
  openOnMount: boolean;
  onSave: (enabled: boolean) => Promise<void>;
}

export interface ISettingsUiHandle {
  unmount: () => void;
}

/** Status bar plus the panel it opens; owns the open/closed and saved state. */
export const SettingsUi: React.FunctionComponent<ISettingsUiProps> = (props: ISettingsUiProps) => {
  const [enabled, setEnabled] = React.useState<boolean>(props.enabled);
  const [open, setOpen] = React.useState<boolean>(props.openOnMount);

  const save = async (next: boolean): Promise<void> => {
    await props.onSave(next);
    setEnabled(next);
  };

  return (
    <>
      {props.canEdit && <SettingsBar enabled={enabled} onOpen={() => setOpen(true)} />}
      <SettingsPanel isOpen={open} enabled={enabled} canEdit={props.canEdit} onSave={save} onDismiss={() => setOpen(false)} />
    </>
  );
};

/**
 * Entry point used by the customizer through a dynamic import, so React and
 * Fluent UI are only downloaded on the page that shows the settings UI.
 */
export function renderSettingsUi(container: HTMLElement, props: ISettingsUiProps): ISettingsUiHandle {
  ReactDom.render(React.createElement(SettingsUi, props), container);
  return {
    unmount: (): void => { ReactDom.unmountComponentAtNode(container); }
  };
}
