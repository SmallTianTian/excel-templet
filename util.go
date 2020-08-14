package xlsxt

import "reflect"

func toStringKeyMap(v interface{}) (map[string]interface{}, error) {
	if v == nil {
		return make(map[string]interface{}), nil
	}
	if skm, ok := v.(map[string]interface{}); ok {
		return skm, nil
	}

	t := reflect.TypeOf(v)
	if t.Kind() != reflect.Map || t.Key().Kind() != reflect.String {
		return nil, NotStringKeyMapValue
	}
	rv := reflect.ValueOf(v)

	skm := make(map[string]interface{})
	iter := rv.MapRange()
	for iter.Next() {
		skm[iter.Key().String()] = iter.Value().Interface()
	}
	return skm, nil
}
