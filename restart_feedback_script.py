# Проверяем, есть ли уже обработанные данные
try:
    existing_df = pd.read_excel('parent_feedback_ollama.xlsx', engine="openpyxl")
    last_processed_index = existing_df['Обратная связь для родителей'].last_valid_index() or -1
except FileNotFoundError:
    last_processed_index = -1

# Основная асинхронная функция
async def main():
    timeout = ClientTimeout(total=6000)  # Увеличиваем timeout до 6000 секунд
    async with aiohttp.ClientSession(timeout=timeout) as session:
        for index, row in df.iterrows():
            if index <= last_processed_index:
                continue  # Пропускаем уже обработанных студентов
            prompt = (f"Составь обратную связь для {row['Имя']}. "
                      f"Идея проекта: {row['Идея проекта']}. "
                      f"Выручка: {row['Выручка с ярмарки']}. "
                      f"Участие: {row['Принял участие в Ярмарке']}.")
            feedback = await fetch_feedback(session, prompt)
            df.at[index, 'Обратная связь для родителей'] = feedback
            # Сохраняем результаты после каждого студента
            df.to_excel('parent_feedback_ollama.xlsx', index=False)
            logging.info(f"Результаты сохранены для студента {index + 1}/{len(df)}: {row['Имя']}")