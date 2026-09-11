import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import * as strings from 'UncheckSiteNavigationApplicationCustomizerStrings';

export interface ISettingsPanelProps {
  isOpen: boolean;
  enabled: boolean;
  canEdit: boolean;
  needsSave: boolean;
  onSave: (enabled: boolean) => Promise<void>;
  onDismiss: () => void;
}

/**
 * Side panel with a single toggle that switches the customizer on or off
 * for the current site. Rendered by the application customizer on Site
 * contents; it is not a web part and does not need a page.
 */
export const SettingsPanel: React.FunctionComponent<ISettingsPanelProps> = (props: ISettingsPanelProps) => {
  const [enabled, setEnabled] = React.useState<boolean>(props.enabled);
  const [saving, setSaving] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const mounted = React.useRef<boolean>(true);

  React.useEffect(() => {
    mounted.current = true;
    return () => { mounted.current = false; };
  }, []);

  React.useEffect(() => {
    if (props.isOpen) {
      setEnabled(props.enabled);
      setError(undefined);
    }
  }, [props.isOpen, props.enabled]);

  const save = (): void => {
    setSaving(true);
    setError(undefined);
    props.onSave(enabled)
      .then(() => { if (mounted.current) { setSaving(false); props.onDismiss(); } })
      .catch((e: Error) => { if (mounted.current) { setSaving(false); setError(e.message || String(e)); } });
  };

  const footer = (): JSX.Element => (
    <Stack horizontal tokens={{ childrenGap: 8 }}>
      <PrimaryButton text={strings.Save} onClick={save} disabled={!props.canEdit || saving || (!props.needsSave && enabled === props.enabled)} data-automation-id="jfdi-unav-save" />
      <DefaultButton text={strings.Cancel} onClick={props.onDismiss} disabled={saving} />
    </Stack>
  );

  return (
    <Panel
      isOpen={props.isOpen}
      onDismiss={props.onDismiss}
      type={PanelType.medium}
      headerText={strings.PanelTitle}
      closeButtonAriaLabel={strings.Cancel}
      isFooterAtBottom={true}
      onRenderFooterContent={footer}
      data-automation-id="jfdi-unav-panel"
    >
      <Stack tokens={{ childrenGap: 16 }}>
        <Text>{strings.PanelDescription}</Text>
        {!props.canEdit && <MessageBar messageBarType={MessageBarType.info}>{strings.ReadOnlyNotice}</MessageBar>}
        <Toggle
          label={strings.ToggleLabel}
          onText={strings.ToggleOn}
          offText={strings.ToggleOff}
          checked={enabled}
          disabled={!props.canEdit || saving}
          onChange={(_e: React.MouseEvent<HTMLElement>, checked?: boolean) => setEnabled(!!checked)}
          data-automation-id="jfdi-unav-toggle"
        />
        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      </Stack>
    </Panel>
  );
};
