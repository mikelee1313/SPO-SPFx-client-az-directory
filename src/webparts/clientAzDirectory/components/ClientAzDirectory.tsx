import * as React from 'react';
import styles from './ClientAzDirectory.module.scss';
import type { IClientAzDirectoryProps } from './IClientAzDirectoryProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface IClientDirectoryItem {
  id: number;
  name: string;
  secondaryText: string;
  url: string;
}

interface IClientAzDirectoryState {
  clients: IClientDirectoryItem[];
  errorMessage: string;
  expandedLetters: string[];
  isLoading: boolean;
  searchText: string;
}

const alphabet: string[] = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
const otherLetter: string = '#';

export default class ClientAzDirectory extends React.Component<IClientAzDirectoryProps, IClientAzDirectoryState> {
  public constructor(props: IClientAzDirectoryProps) {
    super(props);

    this.state = {
      clients: [],
      errorMessage: '',
      expandedLetters: [],
      isLoading: true,
      searchText: ''
    };
  }

  public componentDidMount(): void {
    this._loadClients().catch(error => this._setError(error));
  }

  public componentDidUpdate(previousProps: IClientAzDirectoryProps): void {
    if (
      previousProps.listTitle !== this.props.listTitle ||
      previousProps.nameField !== this.props.nameField ||
      previousProps.secondaryField !== this.props.secondaryField ||
      previousProps.linkField !== this.props.linkField
    ) {
      this._loadClients().catch(error => this._setError(error));
    }
  }

  public render(): React.ReactElement<IClientAzDirectoryProps> {
    const { hasTeamsContext, listTitle } = this.props;
    const { clients, errorMessage, expandedLetters, isLoading, searchText } = this.state;
    const normalizedSearchText: string = searchText.trim().toLocaleLowerCase();
    const filteredClients: IClientDirectoryItem[] = normalizedSearchText
      ? clients.filter(client =>
        client.name.toLocaleLowerCase().indexOf(normalizedSearchText) > -1 ||
        client.secondaryText.toLocaleLowerCase().indexOf(normalizedSearchText) > -1
      )
      : clients;
    const countsByLetter: Map<string, number> = this._getCountsByLetter(clients);
    const visibleLetters: string[] = [...alphabet, otherLetter].filter(letter =>
      (countsByLetter.get(letter) || 0) > 0
    );

    return (
      <section className={`${styles.clientAzDirectory} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.header}>
          <div>
            <h2>Client directory</h2>
            <p>Search by client name or expand a letter to browse the A-Z list.</p>
          </div>
          <div className={styles.totalCount}>{clients.length.toLocaleString()} clients</div>
        </div>

        <label className={styles.searchLabel} htmlFor="clientAzSearch">Search clients</label>
        <input
          className={styles.searchBox}
          id="clientAzSearch"
          onChange={this._onSearchChanged}
          placeholder="Start typing a client name..."
          type="search"
          value={searchText}
        />

        {isLoading && <div className={styles.status}>Loading clients from {listTitle}...</div>}
        {!isLoading && errorMessage && <div className={styles.error}>{errorMessage}</div>}
        {!isLoading && !errorMessage && clients.length === 0 && (
          <div className={styles.status}>No clients were found in {listTitle}.</div>
        )}

        {!isLoading && !errorMessage && clients.length > 0 && (
          <>
            <div className={styles.letterGrid} aria-label="Client counts by first letter">
              {[...alphabet, otherLetter].map(letter => this._renderLetterButton(letter, countsByLetter.get(letter) || 0))}
            </div>

            {normalizedSearchText
              ? this._renderSearchResults(filteredClients)
              : visibleLetters.map(letter => this._renderLetterSection(
                letter,
                clients.filter(client => this._getClientLetter(client.name) === letter),
                expandedLetters.indexOf(letter) > -1
              ))}
          </>
        )}
      </section>
    );
  }

  private async _loadClients(): Promise<void> {
    const { listTitle, nameField, secondaryField, linkField } = this.props;

    if (!listTitle || !nameField) {
      this.setState({
        clients: [],
        errorMessage: 'Configure the SharePoint list name and client name column in the web part properties.',
        isLoading: false
      });
      return;
    }

    this.setState({ errorMessage: '', isLoading: true });

    const fields: string[] = ['Id', nameField, secondaryField, linkField]
      .filter((field, index, allFields) => !!field && allFields.indexOf(field) === index);
    const clients: IClientDirectoryItem[] = [];
    let requestUrl: string = `${this.props.webAbsoluteUrl}/_api/web/lists/getbytitle('${this._escapeODataString(listTitle)}')/items?$select=${fields.join(',')}&$orderby=${nameField} asc&$top=5000`;

    while (requestUrl) {
      const response: SPHttpClientResponse = await this.props.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);

      if (!response.ok) {
        throw new Error(`Unable to load clients from "${listTitle}". SharePoint returned ${response.status} ${response.statusText}.`);
      }

      const page: {
        value: Record<string, unknown>[];
        '@odata.nextLink'?: string;
        'odata.nextLink'?: string;
      } = await response.json();

      page.value.forEach(item => {
        const name: string = this._getTextFieldValue(item, nameField);

        if (name) {
          clients.push({
            id: Number(item.Id),
            name,
            secondaryText: this._getTextFieldValue(item, secondaryField),
            url: this._getLinkFieldValue(item, linkField)
          });
        }
      });

      requestUrl = page['@odata.nextLink'] || page['odata.nextLink'] || '';
    }

    this.setState({
      clients,
      errorMessage: '',
      expandedLetters: [],
      isLoading: false
    });
  }

  private _renderLetterButton(letter: string, count: number): React.ReactElement {
    const isDisabled: boolean = count === 0;

    return (
      <button
        className={`${styles.letterButton} ${isDisabled ? styles.disabledLetter : ''}`}
        disabled={isDisabled}
        key={letter}
        onClick={() => this._toggleLetter(letter)}
        type="button"
      >
        <span>{letter}</span>
        <strong>{count.toLocaleString()}</strong>
      </button>
    );
  }

  private _renderLetterSection(letter: string, clients: IClientDirectoryItem[], isExpanded: boolean): React.ReactElement {
    return (
      <div className={styles.letterSection} key={letter}>
        <button
          className={styles.sectionHeader}
          onClick={() => this._toggleLetter(letter)}
          type="button"
          aria-expanded={isExpanded}
        >
          <span>{letter}</span>
          <span>{clients.length.toLocaleString()} clients</span>
        </button>

        {isExpanded && this._renderClientList(clients)}
      </div>
    );
  }

  private _renderSearchResults(clients: IClientDirectoryItem[]): React.ReactElement {
    return (
      <div className={styles.searchResults}>
        <h3>{clients.length.toLocaleString()} search results</h3>
        {clients.length > 0
          ? this._renderClientList(clients)
          : <div className={styles.status}>No clients match your search.</div>}
      </div>
    );
  }

  private _renderClientList(clients: IClientDirectoryItem[]): React.ReactElement {
    return (
      <ul className={styles.clientList}>
        {clients.map(client => (
          <li key={client.id}>
            {client.url
              ? <a href={client.url} target="_blank" rel="noreferrer">{client.name}</a>
              : <span>{client.name}</span>}
            {client.secondaryText && <small>{client.secondaryText}</small>}
          </li>
        ))}
      </ul>
    );
  }

  private _toggleLetter(letter: string): void {
    this.setState(previousState => {
      const isExpanded: boolean = previousState.expandedLetters.indexOf(letter) > -1;

      return {
        expandedLetters: isExpanded
          ? previousState.expandedLetters.filter(expandedLetter => expandedLetter !== letter)
          : [...previousState.expandedLetters, letter]
      };
    });
  }

  private _onSearchChanged = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ searchText: event.currentTarget.value });
  };

  private _getCountsByLetter(clients: IClientDirectoryItem[]): Map<string, number> {
    const counts: Map<string, number> = new Map<string, number>();

    clients.forEach(client => {
      const letter: string = this._getClientLetter(client.name);
      counts.set(letter, (counts.get(letter) || 0) + 1);
    });

    return counts;
  }

  private _getClientLetter(name: string): string {
    const firstCharacter: string = (name.trim().charAt(0) || otherLetter).toLocaleUpperCase();
    return alphabet.indexOf(firstCharacter) > -1 ? firstCharacter : otherLetter;
  }

  private _getTextFieldValue(item: Record<string, unknown>, fieldName: string): string {
    if (!fieldName || item[fieldName] === undefined || item[fieldName] === null) {
      return '';
    }

    const value: unknown = item[fieldName];

    if (typeof value === 'object' && value !== null && 'Url' in value) {
      return String((value as { Url?: string }).Url || '');
    }

    return String(value);
  }

  private _getLinkFieldValue(item: Record<string, unknown>, fieldName: string): string {
    if (!fieldName || item[fieldName] === undefined || item[fieldName] === null) {
      return '';
    }

    const value: unknown = item[fieldName];

    if (typeof value === 'object' && value !== null && 'Url' in value) {
      return String((value as { Url?: string }).Url || '');
    }

    return String(value);
  }

  private _escapeODataString(value: string): string {
    return value.replace(/'/g, "''");
  }

  private _setError(error: Error): void {
    this.setState({
      clients: [],
      errorMessage: error.message,
      isLoading: false
    });
  }
}
