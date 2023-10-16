export function needsConfiguration(
  preconfiguredListName: string | undefined | null,
  order: string | undefined | null,
  style: string | undefined | null
): boolean {
  return (
    isEmpty(preconfiguredListName) || isEmpty(order) || isEmpty(style)
  );
}

function isEmpty(value: string | undefined | null): boolean {
  return value === undefined || value === null || value.length === 0;
}
