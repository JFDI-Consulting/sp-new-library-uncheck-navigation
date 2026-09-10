import * as React from 'react';
import { Link } from '@fluentui/react/lib/Link';
import { Icon } from '@fluentui/react/lib/Icon';
import * as strings from 'UncheckSiteNavigationApplicationCustomizerStrings';

export interface ISettingsBarProps {
  enabled: boolean;
  onOpen: () => void;
}

const barStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '6px 32px',
  fontSize: 12,
  lineHeight: '16px',
  background: 'var(--neutralLighter, #f3f2f1)',
  color: 'var(--neutralPrimary, #323130)',
  borderBottom: '1px solid var(--neutralLight, #edebe9)'
};

/**
 * One-line status bar shown on Site contents to users who can manage the
 * web. Says what the current default is and offers a link to change it.
 */
export const SettingsBar: React.FunctionComponent<ISettingsBarProps> = (props: ISettingsBarProps) => (
  <div style={barStyle} data-automation-id="jfdi-unav-bar">
    <Icon iconName="Settings" aria-hidden="true" />
    <span role="status">{props.enabled ? strings.StatusEnabled : strings.StatusDisabled}</span>
    <Link onClick={props.onOpen} data-automation-id="jfdi-unav-change">{strings.Change}</Link>
  </div>
);
