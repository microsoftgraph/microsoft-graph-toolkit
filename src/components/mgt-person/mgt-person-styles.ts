import { css } from 'lit-element';
        import { sharedStyles } from '../../styles/shared-styles';
        export const styles = [
          sharedStyles,
          css`
img,
svg,
.person-initials-container {
  width: 100%;
  height: 100%; }

img {
  border: 0;
  border-radius: 50%; }

.person-initials-container {
  height: 100%;
  width: 100%;
  background-color: brown;
  border-radius: 50%;
  font-size: 14px;
  font-weight: 400;
  color: white;
  line-height: 46px;
  vertical-align: baseline;
  text-align: center;
  position: relative;
  font-family: var(--default-font-family); }

.person-initials-container span {
  vertical-align: baseline;
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%); }

`];