import * as React from "react";
import styles from "./BirthdayCalendar.module.scss";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import { Spinner } from "@fluentui/react";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { faker } from "@faker-js/faker";
import * as dayjs from "dayjs";
import "dayjs/locale/nl";

export type BirthdayCalendarProps = {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
};

export default function BirthdayCalendar({
  description,
  isDarkTheme,
  environmentMessage,
  hasTeamsContext,
  userDisplayName,
  context,
}: BirthdayCalendarProps): React.ReactElement<BirthdayCalendarProps> {
  const [users, setUsers] = React.useState<MicrosoftGraph.User[]>([]);

  const getUserProfilePicture = async (
    userId: string
  ): Promise<string | null> => {
    const client = await context.msGraphClientFactory.getClient("3");

    try {
      const response: Blob = await client
        .api(`/users/${userId}/photo/$value`)
        .get();
      const profilePictureUrl = URL.createObjectURL(response);
      return profilePictureUrl;
    } catch (error) {
      return null;
    }
  };

  const fetchUsersWithBirthdays = async (): Promise<void> => {
    let users: MicrosoftGraph.User[] = [];

    const client = await context.msGraphClientFactory.getClient("3");

    let response = await client.api("/users").get();

    while (response["@odata.nextLink"]) {
      response = await client.api(response["@odata.nextLink"]).get();
      users = users.concat(response.value);
    }

    await Promise.all(
      users.map(async (user, index) => {
        const [details, profilePicture] = await Promise.all([
          client
            .api(`/users/${user.id as string}`)
            .select("birthday")
            .get(),
          getUserProfilePicture(user.id as string),
        ]);

        users[index] = {
          ...users[index],
          ...details,
          birthday: faker.date.birthdate().toISOString(),
          photo: {
            id: profilePicture,
          },
        };
      })
    );

    const today = new Date();
    const currentMonth = today.getMonth() + 1;
    const currentDay = today.getDate();

    setUsers(
      users
        .filter(
          ({ birthday }) => birthday && birthday !== "0001-01-01T08:00:00Z"
        )
        .sort((a, b) => {
          const aDate = new Date(a.birthday as string);
          const bDate = new Date(b.birthday as string);

          const aMonth = aDate.getUTCMonth() + 1;
          const aDay = aDate.getUTCDate();

          const bMonth = bDate.getUTCMonth() + 1;
          const bDay = bDate.getUTCDate();

          const aDaysFromToday =
            aMonth < currentMonth ||
            (aMonth === currentMonth && aDay < currentDay)
              ? aMonth * 31 + aDay - currentMonth * 31 - currentDay + 365
              : aMonth * 31 + aDay - currentMonth * 31 - currentDay;

          const bDaysFromToday =
            bMonth < currentMonth ||
            (bMonth === currentMonth && bDay < currentDay)
              ? bMonth * 31 + bDay - currentMonth * 31 - currentDay + 365
              : bMonth * 31 + bDay - currentMonth * 31 - currentDay;

          return aDaysFromToday - bDaysFromToday;
        })
        .slice(0, 5)
    );
  };

  const isSameMonthAndDate = (a: Date, b: Date): boolean =>
    a.getMonth() === b.getMonth() && a.getDate() === b.getDate();

  React.useEffect(() => {
    // eslint-disable-next-line no-void
    void fetchUsersWithBirthdays();
  }, []);

  return (
    <section
      className={`${styles.birthdayCalendar} ${
        hasTeamsContext ? styles.teams : ""
      }`}
    >
      {/* <div className={styles.welcome}>
        <img
          alt=""
          src={
            isDarkTheme
              ? require("../assets/welcome-dark.png")
              : require("../assets/welcome-light.png")
          }
          className={styles.welcomeImage}
        />
        <h2>Well Doen hey, {escape(userDisplayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>
          Web part property value: <strong>{escape(description)}</strong>
        </div>
      </div> */}
      <div className={styles.container}>
        <h2>Aankomende verjaardagen</h2>
        {!users.length ? (
          <Spinner />
        ) : (
          users.map(({ id, displayName, birthday, photo }) => {
            return (
              <div key={id} className={styles.user}>
                {photo?.id && <img src={photo.id} className={styles.image} />}
                <b>{displayName}</b>
                <p>
                  {dayjs(birthday as string)
                    .locale("nl")
                    .format("D MMMM")}
                </p>
                {isSameMonthAndDate(
                  new Date(birthday as string),
                  new Date()
                ) && <h3>🎉</h3>}
              </div>
            );
          })
        )}
      </div>
    </section>
  );
}
