export function mapToRecord<K extends string, V>(map: Map<K, V>): Record<K, V> {
  const obj = {} as Record<K, V>;
  map.forEach((value, key) => {
    obj[key] = value;
  });
  return obj;
}