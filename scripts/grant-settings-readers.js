#!/usr/bin/env node
// Idempotently grant Read on an existing, verified settings list. Never restores inheritance.
const { execFileSync } = require('child_process');
const assert = require('assert');
const COMPONENT_ID = 'd31c6f18-3a0d-462b-b677-c09314fbf3e6';
const [webArgument, ...principalArguments] = process.argv.slice(2);
if (!webArgument || principalArguments.some((id) => !/^[1-9]\d*$/.test(id))) {
  console.error('Usage: node scripts/grant-settings-readers.js <https://tenant/site> [principal-id ...]\nWithout IDs, grants Read to associated Members and Visitors.');
  process.exit(1);
}
const web = new URL(webArgument);
assert.strictEqual(web.protocol, 'https:', 'Use an HTTPS SharePoint web URL');
assert.ok(!web.search && !web.hash && !web.username && !web.password, 'Use a plain web URL');
const webUrl = web.href.replace(/\/$/, '');
const listPath = `${decodeURIComponent(web.pathname).replace(/\/$/, '')}/Lists/JfdiUnavSettings`;
const alias = `@list='${encodeURIComponent(listPath.replace(/'/g, "''"))}'`;
const list = `${webUrl}/_api/web/GetList(@list)`;
function request(url, method = 'get', body) {
  const args = ['request', '--url', url, '--method', method, '--accept', 'application/json;odata=nometadata', '--output', 'json'];
  if (body) args.push('--content-type', 'application/json;odata=nometadata', '--data', JSON.stringify(body));
  const raw = execFileSync('m365', args, { encoding: 'utf8' }).trim();
  return raw ? JSON.parse(raw) : {};
}
try {
  const info = request(`${list}?${alias}&$select=Id,Hidden,HasUniqueRoleAssignments,BaseTemplate,Description`);
  assert.ok(info.Hidden && info.HasUniqueRoleAssignments && info.BaseTemplate === 100 && info.Description === `JFDI Uncheck Navigation settings store:${COMPONENT_ID}`, 'Refusing an unrelated or inherited-permission list');
  const rows = request(`${list}/items?${alias}&$select=Id,JfdiUnavStoreId&$filter=Title%20eq%20'configuration'&$top=2`).value;
  assert.ok(rows && rows.length === 1 && rows[0].JfdiUnavStoreId === COMPONENT_ID, 'Refusing a list without the application configuration marker');
  const groups = principalArguments.length ? undefined : request(`${webUrl}/_api/web?$select=AssociatedMemberGroup/Id,AssociatedVisitorGroup/Id&$expand=AssociatedMemberGroup,AssociatedVisitorGroup`);
  const principals = [...new Set(principalArguments.length ? principalArguments.map(Number) : [groups.AssociatedMemberGroup?.Id, groups.AssociatedVisitorGroup?.Id].filter((id) => id > 0))];
  assert.ok(principals.length, 'No reader principals found; provide SharePoint principal IDs explicitly');
  const role = request(`${webUrl}/_api/web/roledefinitions/getbytype(2)?$select=Id`);
  const assignments = request(`${list}/roleassignments?${alias}&$expand=RoleDefinitionBindings`).value;
  for (const principal of principals) {
    const existing = assignments.find((assignment) => assignment.PrincipalId === principal);
    const readable = existing && existing.RoleDefinitionBindings.some((binding) => [2, 3, 4, 5, 6].includes(binding.RoleTypeKind));
    if (readable) {
      console.log(`Principal ${principal}: existing access retained`);
      continue;
    }
    request(`${list}/roleassignments/addroleassignment(principalid=${principal},roledefid=${role.Id})?${alias}`, 'post');
    const verified = request(`${list}/roleassignments/getbyprincipalid(${principal})/RoleDefinitionBindings?${alias}`).value;
    assert.ok(verified.some((binding) => binding.Id === role.Id), 'Read grant was not confirmed');
    console.log(`Principal ${principal}: Read granted and verified`);
  }
} catch (error) {
  console.error(error.message);
  process.exitCode = 1;
}
