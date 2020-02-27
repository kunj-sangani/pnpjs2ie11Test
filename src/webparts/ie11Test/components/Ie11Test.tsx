import * as React from 'react';
import styles from './Ie11Test.module.scss';
import { IIe11TestProps } from './IIe11TestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";

export default class Ie11Test extends React.Component<IIe11TestProps, {}> {

  public componentDidMount() {
    sp.web.lists.getByTitle('Employee').items.getAll().then((val) => {
      console.log(val);
    }).catch((error) => {
      console.log(error);
    });
  }

  public render(): React.ReactElement<IIe11TestProps> {
    return (
      <div className={styles.ie11Test}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
