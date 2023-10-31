/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { updateMessageContentWithImage } from './updateMessageContentWithImage';

describe('updateMessageContentWithImage', () => {
  it('should replace the image src', () => {
    const content = '<img id="message-id-0" src="https://example.com/old-image.png" />';
    const index = '0';
    const messageId = 'message-id';
    const image = 'https://example.com/new-image.png';
    const expected = '<img id="message-id-0" src="https://example.com/new-image.png" />';
    const actual: string = updateMessageContentWithImage(content, index, messageId, image);
    expect(actual).toEqual(expected);
  });

  it('should inline base64 images', () => {
    const content = `<p>&nbsp;</p>
        <p><img alt="image" id="1683310230186-0" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAMSURBVBhXY7h58yYABRoCjB9qX5UAAAAASUVORK5CYII=" class="image-placeholder" width="440.38461538461536" height="250" itemid="0-wus-d1-6d6b87b78939afbcc919fd49f6696654" style="vertical-align:bottom"></p>
        <p>&nbsp;</p>`;
    const index = '0';
    const messageId = '1683310230186';
    const image =
      'data:image/png;base64,iVBORw0KGAPsxMXFi8TFSlLLFEnu1EmatthLWrVOllZtU6Rl2zaSkblTVm3YJJ99N1O+/n6m/LFoiaxcu17S0neYY9lZmZK2eZNkbNsqMfk50nSvveSww4+S448/SRo2TPLvGhnTX58kKf36yYzplS9mBgzoJ5K6pKjw8u47w1sF72+nH/CuklHvvCOjBoT2Vz2eaH0kVYY8PMqL/yi5sd8Meb2IgvWE9+ve8RsvkPK8kgeMeliGpD5SoQJeqT7c9woEtxVFURSlKigQlDH1vR+hzd6HTaFt1jkv5Uu+U6dxLZgJ98dK3All/+HqEX+QXFLvWjmszjEyssF9cluDh8x245hC9z+2Oc/lhIQhZrFwzuDEc+S6+v8wYbCcUXeof7R8YKVs0qSJvyWy3377FbFkvvTSS3LiiUX7kHH8jjvukDfffLPgPIRqWYkknKeeesoszMn1+eefm/XIkSP9o15enXCCvP766/L++++b5aKLLpJXX31VPvnkE7nqqqv8syKjefPmcvvtt8trr70mkyZNMss999wjzZo1888Icd1118lzzz0nl156qb9HAVdEuuKyMti5M13S07fJ1q2bZdWqlTJ79kz531efy/ezf5KVGTtlU16u/LF6mSxcvlj';

    const expected = `<p>&nbsp;</p>
        <p><img alt="image" id="1683310230186-0" src="data:image/png;base64,iVBORw0KGAPsxMXFi8TFSlLLFEnu1EmatthLWrVOllZtU6Rl2zaSkblTVm3YJJ99N1O+/n6m/LFoiaxcu17S0neYY9lZmZK2eZNkbNsqMfk50nSvveSww4+S448/SRo2TPLvGhnTX58kKf36yYzplS9mBgzoJ5K6pKjw8u47w1sF72+nH/CuklHvvCOjBoT2Vz2eaH0kVYY8PMqL/yi5sd8Meb2IgvWE9+ve8RsvkPK8kgeMeliGpD5SoQJeqT7c9woEtxVFURSlKigQlDH1vR+hzd6HTaFt1jkv5Uu+U6dxLZgJ98dK3All/+HqEX+QXFLvWjmszjEyssF9cluDh8x245hC9z+2Oc/lhIQhZrFwzuDEc+S6+v8wYbCcUXeof7R8YKVs0qSJvyWy3377FbFkvvTSS3LiiUX7kHH8jjvukDfffLPgPIRqWYkknKeeesoszMn1+eefm/XIkSP9o15enXCCvP766/L++++b5aKLLpJXX31VPvnkE7nqqqv8syKjefPmcvvtt8trr70mkyZNMss999wjzZo1888Icd1118lzzz0nl156qb9HAVdEuuKyMti5M13S07fJ1q2bZdWqlTJ79kz531efy/ezf5KVGTtlU16u/LF6mSxcvlj" class="image-placeholder" width="440.38461538461536" height="250" itemid="0-wus-d1-6d6b87b78939afbcc919fd49f6696654" style="vertical-align:bottom"></p>
        <p>&nbsp;</p>`;
    const actual: string = updateMessageContentWithImage(content, index, messageId, image);
    expect(actual).toEqual(expected);
  });
});
