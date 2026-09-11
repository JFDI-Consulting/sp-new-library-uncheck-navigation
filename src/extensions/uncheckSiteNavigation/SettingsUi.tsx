import * as React from 'react';
import * as ReactDom from 'react-dom';
import { SettingsBar } from './SettingsBar';
import { SettingsPanel } from './SettingsPanel';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

export interface ISettingsUiProps {
  enabled: boolean;
  canEdit: boolean;
  showBar: boolean;
  loadError?: string;
  onOpen: () => Promise<{ enabled: boolean; needsSave: boolean }>;
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
  const [needsSave, setNeedsSave] = React.useState<boolean>(false);
  const [open, setOpen] = React.useState<boolean>(false);

  const [error, setError] = React.useState<string | undefined>(props.loadError);
  const [loading, setLoading] = React.useState<boolean>(false);
  const mounted = React.useRef<boolean>(true);
  const openPanel = (): void => {
    setLoading(true);
    setError(undefined);
    props.onOpen().then((current: { enabled: boolean; needsSave: boolean }) => {
      if (mounted.current) { setEnabled(current.enabled); setNeedsSave(current.needsSave); setOpen(true); setLoading(false); }
    }).catch((e: Error) => {
      if (mounted.current) { setError(e.message); setLoading(false); }
    });
  };
  React.useEffect(() => {
    mounted.current = true;
    if (props.openOnMount) { openPanel(); }
    return () => { mounted.current = false; };
  }, []);

  const save = async (next: boolean): Promise<void> => {
    await props.onSave(next);
    if (mounted.current) { setEnabled(next); setNeedsSave(false); }
  };

  return (
    <>
      {props.showBar && <SettingsBar enabled={enabled} unavailable={!!error} loading={loading} onOpen={openPanel} />}
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      <SettingsPanel needsSave={needsSave} isOpen={open} enabled={enabled} canEdit={props.canEdit} onSave={save} onDismiss={() => setOpen(false)} />
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
